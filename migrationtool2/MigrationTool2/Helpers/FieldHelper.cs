using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;

namespace MigrationTool2.Helpers;

public class FieldHelper
{
    private readonly Models.OptionsConfig _options;
    private readonly Dictionary<string, Dictionary<int, int>> _globalIdMapping;
    private readonly Dictionary<string, string> _listNameCache = new();
    private static User? _currentUser;
    private readonly Services.AuthService? _authService;
    private readonly string _destSiteUrl;
    private readonly HttpClient _httpClient = new();
    private string? _cachedToken;
    private int? _tzBiasMinutes;

    private static readonly HashSet<string> SystemFields = new(StringComparer.OrdinalIgnoreCase)
    {
        "ID", "ContentTypeId", "ContentType", "UniqueId", "GUID",
        "WorkflowInstanceID", "WorkflowVersion", "FileRef", "FileDirRef",
        "FileLeafRef", "File_x0020_Type", "FSObjType", "PermMask",
        "IsCheckedoutToLocal", "CheckedOutUserId", "CheckoutUser",
        "CheckedOutTitle", "IsCurrentVersion", "AppCreatedBy",
        "AppModifiedBy", "SMTotalSize", "SMLastModifiedDate",
        "_UIVersion", "_UIVersionString", "ParentLeaf", "ParentUniqueId",
        "ScopeId", "DocIcon", "HTML_x0020_File_x0020_Type",
        "Edit", "ServerUrl", "EncodedAbsUrl", "BaseName",
        "MetaInfo", "_Level", "_IsCurrentVersion", "owshiddenversion",
        "InstanceID", "Order", "WikiField",
        "Attachments", "_ColorTag", "ComplianceAssetId"
    };

    public FieldHelper(
        Models.OptionsConfig options,
        Dictionary<string, Dictionary<int, int>> globalIdMapping,
        Services.AuthService? authService = null,
        string destSiteUrl = "")
    {
        _options = options;
        _globalIdMapping = globalIdMapping;
        _authService = authService;
        _destSiteUrl = destSiteUrl.TrimEnd('/');
    }

    public (ListItem createdItem, int destId) CreateAndCopyFields(
        List destList, ListItem sourceItem, ClientContext sourceCtx, ClientContext destCtx,
        Dictionary<string, Field> destFields, string sourceListName,
        string destListName,
        ListItem? existingItem = null)
    {
        if (_authService != null && !string.IsNullOrEmpty(_destSiteUrl))
        {
            try
            {
                var (restId, _) = RestCreateAndCopyFields(
                    destListName, sourceItem, sourceCtx, destCtx,
                    destFields, sourceListName, existingItem).GetAwaiter().GetResult();
                return (existingItem ?? destList.GetItemById(restId), restId);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"  [REST] Error: {ex.Message}");
                throw;
            }
        }

        var item = existingItem;
        int destId = 0;

        if (existingItem != null)
        {
            destId = existingItem.Id;
        }
        else
        {
            var createInfo = new ListItemCreationInformation();
            item = destList.AddItem(createInfo);
        }

        foreach (var (fieldName, destField) in destFields)
        {
            try
            {
                if (SystemFields.Contains(fieldName))
                    continue;

                if (IsTaxonomyField(destField))
                {
                    SetTaxonomyFieldInMemory(sourceItem, item, destField, destCtx, sourceListName);
                    continue;
                }

                var value = ReadFieldValue(sourceItem, fieldName);
                if (ShouldSkipField(fieldName, value))
                    continue;

                var converted = ConvertFieldValue(fieldName, value, destField, sourceListName, destCtx, sourceCtx);
                if (converted != null)
                    item[fieldName] = converted;
            }
            catch { }
        }

        try
        {
            var title = ReadFieldValue(sourceItem, "Title");
            if (title != null)
                item.ParseAndSetFieldValue("Title", title.ToString());
        }
        catch { }

        SetSystemFields(item, sourceItem, sourceCtx, destCtx);

        if (existingItem != null)
        {
            item.SystemUpdate();
            destCtx.ExecuteQuery();
        }
        else
        {
            item.Update();
            destCtx.ExecuteQuery();
            destId = item.Id;
            try
            {
                var title = ReadFieldValue(sourceItem, "Title");
                if (title != null)
                    item.ParseAndSetFieldValue("Title", title.ToString());

                if (_options.PreserveModifiedBy)
                {
                    var e = ReadFieldValue(sourceItem, "Editor");
                    if (e != null)
                    {
                        var email = ResolveUserEmail(e, sourceCtx);
                        if (!string.IsNullOrEmpty(email))
                        {
                            var user = ResolveUser(destCtx, email);
                            if (user != null)
                                item["Editor"] = new FieldUserValue { LookupId = user.Id };
                        }
                    }
                }
                if (_options.PreserveModified)
                {
                    var m = ReadFieldValue(sourceItem, "Modified");
                    if (m is DateTime dt) item["Modified"] = dt;
                    else if (m is string s && DateTime.TryParse(s, out var p)) item["Modified"] = p;
                }
                item.SystemUpdate();
                destCtx.ExecuteQuery();
            }
            catch { }
        }

        return (item, destId);
    }

    public void SystemUpdateExistingItem(
        ListItem destItem, ListItem sourceItem, ClientContext sourceCtx, ClientContext destCtx)
    {
        SetSystemFields(destItem, sourceItem, sourceCtx, destCtx);
        try
        {
            destItem.SystemUpdate();
            destCtx.ExecuteQuery();
        }
        catch
        {
            destItem.Update();
            destCtx.ExecuteQuery();
        }
    }

    // ────────────── REST API Helpers ──────────────

    private async Task<string> GetRestTokenAsync()
    {
        if (_authService == null) throw new InvalidOperationException("AuthService required");
        return await _authService.GetAccessTokenAsync(_destSiteUrl);
    }

    private async Task<(int destId, ListItem? item)> RestCreateAndCopyFields(
        string destListName, ListItem sourceItem, ClientContext sourceCtx, ClientContext destCtx,
        Dictionary<string, Field> destFields, string sourceListName,
        ListItem? existingItem = null)
    {
        var token = await GetRestTokenAsync();
        var baseUrl = _destSiteUrl;
        var listNameEscaped = destListName.Replace("'", "''");

        // Collect field form values
        var formValues = new List<Dictionary<string, object>>();

        foreach (var (fieldName, destField) in destFields)
        {
            try
            {
                if (SystemFields.Contains(fieldName)) continue;
                var value = ReadFieldValue(sourceItem, fieldName);
                if (ShouldSkipField(fieldName, value)) continue;

                var strVal = FieldValueToRestString(fieldName, value, destField, sourceListName, destCtx, sourceCtx);
                if (strVal != null)
                    formValues.Add(new Dictionary<string, object> { ["FieldName"] = fieldName, ["FieldValue"] = strVal });
            }
            catch { }
        }

        // Add Title
        try
        {
            var title = ReadFieldValue(sourceItem, "Title");
            if (title != null)
                formValues.Add(new Dictionary<string, object> { ["FieldName"] = "Title", ["FieldValue"] = title.ToString()! });
        }
        catch { }

        if (existingItem != null)
        {
            // Update existing item
            var ok = await RestValidateUpdate(destListName, existingItem.Id, formValues, false, token, baseUrl);
            if (!ok) throw new Exception("REST ValidateUpdateListItem failed for existing item");

            // Update system fields
            var sysVals = BuildSystemFieldFormValues(sourceItem, sourceCtx, destCtx);
            if (sysVals.Count > 0)
                await RestValidateUpdate(destListName, existingItem.Id, sysVals, true, token, baseUrl);

            return (existingItem.Id, existingItem);
        }
        else
        {
            // Create item via POST with all field values as JSON properties
            var jsonObj = new Dictionary<string, object>();
            foreach (var fv in formValues)
            {
                if (fv.TryGetValue("FieldName", out var nameObj) && fv.TryGetValue("FieldValue", out var valObj))
                    jsonObj[nameObj.ToString()!] = valObj.ToString() ?? "";
            }

            var json = JsonSerializer.Serialize(jsonObj);

            var url = $"{baseUrl}/_api/web/lists/getbytitle('{listNameEscaped}')/items";

            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            _httpClient.DefaultRequestHeaders.Remove("Accept");
            _httpClient.DefaultRequestHeaders.Add("Accept", "application/json;odata=verbose");

            var content = new StringContent(json, Encoding.UTF8, "application/json");
            var response = await _httpClient.PostAsync(url, content);
            var responseBody = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
                throw new Exception($"REST create item failed ({response.StatusCode}): {responseBody}");

            var doc = JsonDocument.Parse(responseBody);
            var newId = doc.RootElement.GetProperty("d").GetProperty("Id").GetInt32();

            // Second pass: system fields via ValidateUpdateListItem
            var sysValsNew = BuildSystemFieldFormValues(sourceItem, sourceCtx, destCtx);
            if (sysValsNew.Count > 0)
                await RestValidateUpdate(destListName, newId, sysValsNew, true, token, baseUrl);

            return (newId, null);
        }
    }

    private async Task<bool> RestValidateUpdate(string destListName, int itemId,
        List<Dictionary<string, object>> formValues, bool bNewDocumentUpdate,
        string token, string baseUrl)
    {
        var listNameEscaped = destListName.Replace("'", "''");
        var url = $"{baseUrl}/_api/web/lists/getbytitle('{listNameEscaped}')/items({itemId})/ValidateUpdateListItem";

        var body = JsonSerializer.Serialize(new
        {
            formValues,
            bNewDocumentUpdate
        });

        _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
        _httpClient.DefaultRequestHeaders.Remove("Accept");
        _httpClient.DefaultRequestHeaders.Add("Accept", "application/json;odata=verbose");

        var content = new StringContent(body, Encoding.UTF8, "application/json");
        var response = await _httpClient.PostAsync(url, content);
        var responseBody = await response.Content.ReadAsStringAsync();

        if (!response.IsSuccessStatusCode)
        {
            Console.Error.WriteLine($"  [REST] ValidateUpdateListItem failed ({response.StatusCode}): {responseBody}");
            return false;
        }

        // Log per-field errors from ValidateUpdateListItem
        try
        {
            var doc = JsonDocument.Parse(responseBody);
            var results = doc.RootElement
                .GetProperty("d")
                .GetProperty("ValidateUpdateListItem")
                .GetProperty("results")
                .EnumerateArray();
            foreach (var r in results)
            {
                if (r.GetProperty("HasException").GetBoolean())
                {
                    var fieldName = r.GetProperty("FieldName").GetString();
                    var errorMsg = r.GetProperty("ErrorMessage").GetString();
                    Console.Error.WriteLine($"  [REST] ValidateUpdate field '{fieldName}' error: {errorMsg}");
                }
            }
        }
        catch { }

        return true;
    }

    private string? FieldValueToRestString(string fieldName, object value, Field destField,
        string sourceListName, ClientContext destCtx, ClientContext? sourceCtx)
    {
        if (value == null) return null;
        if (value is string s) return s;
        if (value is DateTime dt) return dt.ToString("o");
        if (value is bool b) return b ? "1" : "0";
        if (value is int i) return i.ToString();
        if (value is FieldUserValue fuv)
        {
            if (!string.IsNullOrEmpty(fuv.Email)) return fuv.Email;
            if (fuv.LookupId > 0)
            {
                var email = ResolveUserEmail(fuv, sourceCtx);
                if (!string.IsNullOrEmpty(email)) return email;
            }
            return fuv.LookupValue;
        }
        if (value is FieldLookupValue flv)
            return flv.LookupValue;
        if (value is TaxonomyFieldValue tfv)
            return $"{tfv.TermGuid};#{tfv.Label}";
        return value.ToString();
    }

    private List<Dictionary<string, object>> BuildSystemFieldFormValues(
        ListItem sourceItem, ClientContext sourceCtx, ClientContext destCtx)
    {
        var vals = new List<Dictionary<string, object>>();
        try
        {
            if (_options.PreserveCreatedBy)
            {
                var a = ReadFieldValue(sourceItem, "Author");
                if (a is FieldUserValue fuv)
                {
                    var email = ResolveUserEmail(fuv, sourceCtx);
                    if (!string.IsNullOrEmpty(email))
                        vals.Add(new() { ["FieldName"] = "Author",
                            ["FieldValue"] = $"[{{'Key':'i:0#.f|membership|{email}'}}]" });
                }
            }
        }
        catch { }
        try
        {
            if (_options.PreserveModifiedBy)
            {
                var e = ReadFieldValue(sourceItem, "Editor");
                if (e is FieldUserValue fuv)
                {
                    var email = ResolveUserEmail(fuv, sourceCtx);
                    if (!string.IsNullOrEmpty(email))
                        vals.Add(new() { ["FieldName"] = "Editor",
                            ["FieldValue"] = $"[{{'Key':'i:0#.f|membership|{email}'}}]" });
                }
            }
        }
        catch { }
        try
        {
            if (_options.PreserveCreated)
            {
                var c = ReadFieldValue(sourceItem, "Created");
                if (c is DateTime dt)
                    vals.Add(new() { ["FieldName"] = "Created",
                        ["FieldValue"] = UtcToSiteLocalTime(dt, destCtx) });
                else if (c is string s && DateTime.TryParse(s, out var p))
                    vals.Add(new() { ["FieldName"] = "Created",
                        ["FieldValue"] = UtcToSiteLocalTime(p, destCtx) });
            }
        }
        catch { }
        try
        {
            if (_options.PreserveModified)
            {
                var m = ReadFieldValue(sourceItem, "Modified");
                if (m is DateTime dt)
                    vals.Add(new() { ["FieldName"] = "Modified",
                        ["FieldValue"] = UtcToSiteLocalTime(dt, destCtx) });
                else if (m is string s && DateTime.TryParse(s, out var p))
                    vals.Add(new() { ["FieldName"] = "Modified",
                        ["FieldValue"] = UtcToSiteLocalTime(p, destCtx) });
            }
        }
        catch { }
        return vals;
    }

    private int _tzBias = int.MinValue;
    private int _tzDaylightBias;

    private (int standardBias, int daylightBias) GetSiteTimeZoneInfo(ClientContext destCtx)
    {
        if (_tzBias != int.MinValue) return (_tzBias, _tzDaylightBias);

        try
        {
            var web = destCtx.Web;
            destCtx.Load(web.RegionalSettings);
            destCtx.ExecuteQuery();
            var tz = web.RegionalSettings.TimeZone;
            destCtx.Load(tz, t => t.Information);
            destCtx.ExecuteQuery();
            _tzBias = tz.Information.Bias;
            _tzDaylightBias = tz.Information.DaylightBias;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"  [TZ] Failed: {ex.Message}");
            _tzBias = 0;
            _tzDaylightBias = 0;
        }
        return (_tzBias, _tzDaylightBias);
    }

    private string UtcToSiteLocalTime(DateTime dt, ClientContext destCtx)
    {
        var (bias, daylightBias) = GetSiteTimeZoneInfo(destCtx);
        var utcDate = dt.Kind == DateTimeKind.Utc ? dt : dt.ToUniversalTime();
        // Standard bias: minutes to add to local time to get UTC
        // DaylightBias: adjustment during DST (typically negative, e.g., -60)
        // Effective bias = bias + daylightBias when DST is active
        var effectiveBias = bias;
        if (daylightBias != 0 && IsDaylightSaving(utcDate, bias, daylightBias))
            effectiveBias = bias + daylightBias;
        // UTC = Local + effectiveBias  =>  Local = UTC - effectiveBias
        var siteLocal = utcDate.AddMinutes(-effectiveBias);
        return siteLocal.ToString("MM/dd/yyyy HH:mm:ss");
    }

    private static bool IsDaylightSaving(DateTime utcDate, int standardBias, int daylightBias)
    {
        // If standard bias is larger than daylight-adjusted bias, DST reduces the offset (e.g., PST 480 → PDT 420)
        // If standard bias is smaller, DST increases the offset
        // Simple heuristic: check what the local time WOULD be with standard vs daylight
        var stdLocal = utcDate.AddMinutes(-standardBias);
        var dstLocal = utcDate.AddMinutes(-(standardBias + daylightBias));
        // On DST transition dates, check which one is in DST (later local time = DST in northern hemisphere)
        // This is simplified; a full implementation would use transition rules
        // For most practical purposes, if daylightBias < 0 (US/Europe), DST is in effect during spring/summer
        if (utcDate.Month >= 3 && utcDate.Month <= 10)
            return daylightBias < 0;
        return daylightBias > 0;
    }

    private void AddSystemFieldsToJson(Dictionary<string, object> json,
        ListItem sourceItem, ClientContext sourceCtx, ClientContext destCtx)
    {
        try
        {
            if (_options.PreserveCreatedBy)
            {
                var a = ReadFieldValue(sourceItem, "Author");
                if (a is FieldUserValue fuv)
                {
                    var email = ResolveUserEmail(fuv, sourceCtx);
                    if (!string.IsNullOrEmpty(email))
                        json["AuthorId"] = ResolveUserLookupId(destCtx, email);
                }
            }
        }
        catch { }
        try
        {
            if (_options.PreserveModifiedBy)
            {
                var e = ReadFieldValue(sourceItem, "Editor");
                if (e is FieldUserValue fuv)
                {
                    var email = ResolveUserEmail(fuv, sourceCtx);
                    if (!string.IsNullOrEmpty(email))
                        json["EditorId"] = ResolveUserLookupId(destCtx, email);
                }
            }
        }
        catch { }
        try
        {
            if (_options.PreserveCreated)
            {
                var c = ReadFieldValue(sourceItem, "Created");
                if (c is DateTime dt) json["Created"] = dt.ToString("o");
                else if (c is string s && DateTime.TryParse(s, out var p)) json["Created"] = p.ToString("o");
            }
        }
        catch { }
        try
        {
            if (_options.PreserveModified)
            {
                var m = ReadFieldValue(sourceItem, "Modified");
                if (m is DateTime dt) json["Modified"] = dt.ToString("o");
                else if (m is string s && DateTime.TryParse(s, out var p)) json["Modified"] = p.ToString("o");
            }
        }
        catch { }
    }

    private int ResolveUserLookupId(ClientContext ctx, string email)
    {
        try
        {
            var user = ctx.Web.EnsureUser(email);
            ctx.Load(user);
            ctx.ExecuteQuery();
            return user.Id;
        }
        catch { return -1; }
    }

    private void SetSystemFields(ListItem item, ListItem sourceItem, ClientContext sourceCtx, ClientContext destCtx)
    {
        try
        {
            if (_options.PreserveCreatedBy)
            {
                var a = ReadFieldValue(sourceItem, "Author");
                if (a != null)
                {
                    var email = ResolveUserEmail(a, sourceCtx);
                    if (!string.IsNullOrEmpty(email))
                    {
                        var user = ResolveUser(destCtx, email);
                        if (user != null)
                            item["Author"] = new FieldUserValue { LookupId = user.Id };
                    }
                }
            }
        }
        catch { }

        try
        {
            if (_options.PreserveCreated)
            {
                var c = ReadFieldValue(sourceItem, "Created");
                if (c is DateTime dt) item["Created"] = dt;
                else if (c is string s && DateTime.TryParse(s, out var p)) item["Created"] = p;
            }
        }
        catch { }

        try
        {
            if (_options.PreserveModifiedBy)
            {
                var e = ReadFieldValue(sourceItem, "Editor");
                if (e != null)
                {
                    var email = ResolveUserEmail(e, sourceCtx);
                    if (!string.IsNullOrEmpty(email))
                    {
                        var user = ResolveUser(destCtx, email);
                        if (user != null)
                            item["Editor"] = new FieldUserValue { LookupId = user.Id };
                    }
                }
            }
        }
        catch { }

        try
        {
            if (_options.PreserveModified)
            {
                var m = ReadFieldValue(sourceItem, "Modified");
                if (m is DateTime dt) item["Modified"] = dt;
                else if (m is string s && DateTime.TryParse(s, out var p)) item["Modified"] = p;
            }
        }
        catch { }
    }

    private string? ResolveUserEmail(object value, ClientContext sourceCtx)
    {
        if (value is FieldUserValue fuv)
        {
            if (!string.IsNullOrEmpty(fuv.Email))
                return fuv.Email;
            if (fuv.LookupId > 0)
            {
                try
                {
                    var user = sourceCtx.Web.SiteUsers.GetById(fuv.LookupId);
                    sourceCtx.Load(user, u => u.Email);
                    sourceCtx.ExecuteQuery();
                    return user.Email;
                }
                catch { }
            }
        }
        return null;
    }

    private static object? ReadFieldValue(ListItem item, string fieldName)
    {
        try
        {
            var val = item[fieldName];
            if (val != null) return val;
        }
        catch { }

        try
        {
            var fv = item.FieldValuesAsText;
            if (fv == null) return null;
            var val = fv[fieldName];
            if (val != null && !string.IsNullOrEmpty(val.ToString())) return val;
        }
        catch { }

        return null;
    }

    private bool ShouldSkipField(string fieldName, object? value)
    {
        if (string.IsNullOrEmpty(fieldName) || SystemFields.Contains(fieldName))
            return true;
        if (value == null)
            return true;
        return false;
    }

    private object? ConvertFieldValue(string fieldName, object value, Field destField, string sourceListName, ClientContext destCtx, ClientContext? sourceCtx = null)
    {
        if (value == null) return null;

        if (destField.FieldTypeKind == FieldType.Lookup)
        {
            var lookupField = destField as FieldLookup;
            return ConvertLookupValue(value, lookupField, sourceListName);
        }

        if (destField.FieldTypeKind == FieldType.User)
        {
            return ConvertUserValue(value, destCtx, sourceCtx);
        }

        if (destField.FieldTypeKind == FieldType.MultiChoice)
        {
            if (value is string s)
                return s.Split(";#", StringSplitOptions.RemoveEmptyEntries);
            return value;
        }

        if (destField.FieldTypeKind == FieldType.URL)
        {
            if (value is string s2)
            {
                var parts = s2.Split(", ", 2);
                return new FieldUrlValue { Url = parts[0], Description = parts.Length > 1 ? parts[1] : parts[0] };
            }
            return value;
        }

        if (destField.FieldTypeKind == FieldType.DateTime)
        {
            if (value is DateTime dt) return dt;
            if (value is string dateStr && DateTime.TryParse(dateStr, out var parsed)) return parsed;
            return value;
        }

        if (destField.FieldTypeKind == FieldType.Number || destField.FieldTypeKind == FieldType.Integer || destField.FieldTypeKind == FieldType.Currency)
        {
            if (value is double d) return d;
            if (value is int i) return (double)i;
            if (value is string ns && double.TryParse(ns, out var n)) return n;
            return null;
        }

        if (destField.FieldTypeKind == FieldType.Boolean)
        {
            if (value is bool b) return b;
            if (value is string bs && (bs == "Yes" || bs == "1" || bs == "true")) return true;
            if (value is string bs2 && (bs2 == "No" || bs2 == "0" || bs2 == "false")) return false;
            return null;
        }

        return value;
    }

    private void SetTaxonomyFieldInMemory(ListItem sourceItem, ListItem destItem, Field destField, ClientContext destCtx, string sourceListName)
    {
        try
        {
            var taxField = destCtx.CastTo<TaxonomyField>(destField);
            destCtx.Load(taxField, f => f.Id, f => f.InternalName, f => f.TextField);
            destCtx.ExecuteQuery();

            Guid textFieldId = taxField.TextField;
            string textFieldName = taxField.InternalName + "_0";
            var sourceTextValue = sourceItem[textFieldName]?.ToString();
            if (string.IsNullOrEmpty(sourceTextValue))
                sourceTextValue = sourceItem[taxField.InternalName]?.ToString();
            if (string.IsNullOrEmpty(sourceTextValue)) return;

            var labels = sourceTextValue
                .Split(";#", StringSplitOptions.RemoveEmptyEntries)
                .Select(l => l.Trim())
                .Where(l => !string.IsNullOrEmpty(l))
                .Distinct()
                .ToArray();

            if (labels.Length == 0) return;

            var session = TaxonomySession.GetTaxonomySession(destCtx);
            var termStore = session.GetDefaultKeywordsTermStore();
            destCtx.Load(termStore);
            destCtx.ExecuteQuery();

            var termSet = termStore.GetTermSet(taxField.TermSetId);
            destCtx.Load(termSet, ts => ts.Terms.Include(t => t.Id, t => t.Name));
            destCtx.ExecuteQuery();

            var allTerms = termSet.Terms.ToList();
            var matchedTerms = allTerms
                .Where(t => labels.Any(l => string.Equals(t.Name, l, StringComparison.OrdinalIgnoreCase)))
                .ToList();

            if (matchedTerms.Count == 0)
            {
                Console.WriteLine($"  [WARN] No taxonomy terms matched in destination for '{string.Join(", ", labels)}'");
                return;
            }

            foreach (var unmatched in labels.Where(l => !allTerms.Any(t => string.Equals(t.Name, l, StringComparison.OrdinalIgnoreCase))))
            {
                Console.WriteLine($"  [WARN] Taxonomy term '{unmatched}' not found on destination. Skipping.");
            }

            try
            {
                var value = CreateTaxonomyFieldValue(matchedTerms[0]);
                taxField.SetFieldValueByValue(destItem, value);
            }
            catch
            {
                var labelsStr = string.Join(";", matchedTerms.Select(t => t.Name));
                destItem[textFieldName] = labelsStr;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  [WARN] Failed to copy taxonomy field '{destField.InternalName}': {ex.Message}");
        }
    }

    private object? ConvertLookupValue(object value, FieldLookup? lookupField, string sourceListName)
    {
        if (lookupField == null) return null;

        var targetListName = GetListNameFromLookup(lookupField);
        if (targetListName == null) return null;

        if (value is FieldLookupValue singleLookup)
        {
            var destId = _globalIdMapping.GetValueOrDefault(targetListName, new())
                .GetValueOrDefault(singleLookup.LookupId, -1);
            if (destId <= 0) return null;
            return new FieldLookupValue { LookupId = destId };
        }

        if (value is FieldLookupValue[] multiLookup)
        {
            var mapping = _globalIdMapping.GetValueOrDefault(targetListName, new());
            var resolved = multiLookup
                .Select(l => mapping.GetValueOrDefault(l.LookupId, -1))
                .Where(id => id > 0)
                .Select(id => new FieldLookupValue { LookupId = id })
                .ToArray();
            return resolved.Length > 0 ? resolved : null;
        }

        if (value is string s && s.Contains(";#"))
        {
            var parts = s.Split(";#", StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length >= 1 && int.TryParse(parts[0], out var lookupId))
            {
                var destId = _globalIdMapping.GetValueOrDefault(targetListName, new())
                    .GetValueOrDefault(lookupId, -1);
                if (destId > 0)
                    return new FieldLookupValue { LookupId = destId };
            }
        }

        return null;
    }

    private object? ConvertUserValue(object value, ClientContext ctx, ClientContext? sourceCtx = null)
    {
        if (value is FieldUserValue userValue)
        {
            var email = userValue.Email;
            if (string.IsNullOrEmpty(email) && sourceCtx != null)
                email = ResolveUserEmail(userValue, sourceCtx);

            if (!string.IsNullOrEmpty(email))
            {
                var user = ResolveUser(ctx, email);
                if (user != null) return new FieldUserValue { LookupId = user.Id };
            }
            if (userValue.LookupId > 0)
                return userValue.LookupId;
        }

        if (value is FieldUserValue[] userMulti)
        {
            var resolved = new List<FieldUserValue>();
            foreach (var uv in userMulti)
            {
                var email = uv.Email;
                if (string.IsNullOrEmpty(email) && sourceCtx != null)
                    email = ResolveUserEmail(uv, sourceCtx);

                if (!string.IsNullOrEmpty(email))
                {
                    var user = ResolveUser(ctx, email);
                    if (user != null) resolved.Add(new FieldUserValue { LookupId = user.Id });
                }
            }
            return resolved.Count > 0 ? resolved.ToArray() : null;
        }

        if (value is string s && !string.IsNullOrEmpty(s))
        {
            var email = s.Contains('|') ? s.Split('|').LastOrDefault() : s;
            email = email?.Trim();
            if (!string.IsNullOrEmpty(email) && email.Contains('@'))
            {
                var user = ResolveUser(ctx, email);
                if (user != null) return new FieldUserValue { LookupId = user.Id };
            }
        }

        return null;
    }

    private User? ResolveUser(ClientContext ctx, object value)
    {
        var email = value switch
        {
            FieldUserValue fuv => !string.IsNullOrEmpty(fuv.Email) ? fuv.Email : fuv.LookupValue,
            string s when s.Contains('|') => s.Split('|').LastOrDefault()?.Trim(),
            string s => s,
            _ => null
        };

        if (string.IsNullOrEmpty(email) || !email.Contains('@'))
            return null;

        try
        {
            var user = ctx.Web.EnsureUser(email);
            ctx.Load(user);
            ctx.ExecuteQuery();
            return user;
        }
        catch
        {
            try
            {
                if (_currentUser == null)
                {
                    _currentUser = ctx.Web.CurrentUser;
                    ctx.Load(_currentUser);
                    ctx.ExecuteQuery();
                }
                return _currentUser;
            }
            catch
            {
                return null;
            }
        }
    }

    private static bool IsMultiValueTaxonomyField(TaxonomyField taxField)
    {
        try
        {
            var fieldType = taxField.SchemaXml ?? "";
            return fieldType.Contains("MultiValue=\"TRUE\"", StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    private static TaxonomyFieldValue CreateTaxonomyFieldValue(Term term)
    {
        var val = new TaxonomyFieldValue();
        val.Label = term.Name;
        val.TermGuid = term.Id.ToString();
        val.WssId = -1;
        return val;
    }

    public static bool IsTaxonomyField(Field field)
    {
        try
        {
            return field is TaxonomyField;
        }
        catch
        {
            return false;
        }
    }

    private string? GetListNameFromLookup(FieldLookup lookupField)
    {
        var lookupListId = lookupField.LookupList;
        if (string.IsNullOrEmpty(lookupListId)) return null;

        if (_listNameCache.TryGetValue(lookupListId, out var name))
            return name;

        return lookupField.LookupList;
    }

    public void SetListNameCache(string listId, string listName)
    {
        _listNameCache[listId] = listName;
    }

    public static List<Field> GetCopyableFields(List sourceList, ClientContext ctx)
    {
        var fields = sourceList.Fields;
        ctx.Load(fields);
        ctx.ExecuteQuery();

        return fields
            .Where(f => !f.Hidden && !f.ReadOnlyField && !SystemFields.Contains(f.InternalName))
            .Where(f => f.FieldTypeKind != FieldType.Computed)
            .Where(f => !f.InternalName.StartsWith("_"))
            .ToList();
    }
}
