package com.ahsay.afc.cloud.office365.sharepoint;

import com.independentsoft.share.queryoptions.*;

import java.util.ArrayList;

/*
 * Copyright (c) 2017 Ahsay Systems Corporation Limited. All Rights Reserved.
 *
 * Description: SharePoint Query Utils
 *
 * Date        Task  Author           Changes
 * 2017-02-21 15373  marcus.leung     Created
 * 2017-11-20 16141  tony.kung        Added list field filter
 * 2018-10-16 22256  felix.chou       Modify groupOption to get group owner info
 * 2018-10-31 22494  felix.chou       Add support to Calendar list content
 * 2019-01-10 21984  fellx.chou       Add to get list item shareInfo option
 * 2019-02-14 23122  felix.chou       Added query method to be used in getListItems
 * 2019-02-21 23203  felix.chou       Added ListData query
 * 2019-03-20 23322  felix.chou       Added restrictions for ListItem, Field and ContentType
 * 2019-04-03 23474  felix.chou       Modified methods to support query options in listRestorer
 * 2019-04-18 22014  felix.chou       Added site properties query
 * 2019-05-06 23535  nicholas.leung   Support item versions
 * 2019-05-15 23699  nicholas.leung   Add support to External Data
 * 2019-05-31 23597  terry.li         AppRequest list does not support AttachmentFiles but it has attachments enabled
 * 2019-07-15 24258  nicholas.leung   Support to load complete info only when required
 * 2019-09-19 25054  terry.li         Load all info on user information list
 * 2019-10-19 23626  jefferson.brigino Added new fields Title, SiteLogoUrl, QuickLaunchEnabled, TreeViewEnabled to support restore
 */
public class SharePointQueryUtils {

    // [Start] 24258: Support to load complete info only when required
    public static final int MODE_OVERVIEW = 0;
    public static final int MODE_DETAIL = 1;
    // [End] 24258

    // [Start] 23474: Support to use top option alone
    public static ArrayList<IQueryOption> getPaginationOption(int iOffset, int iSize) {
        ArrayList<IQueryOption> queryOptions = new ArrayList<IQueryOption>();

        queryOptions.add(new Skip(iOffset));
        queryOptions.addAll(getTopOption(iSize));
        return queryOptions;
    }

    public static ArrayList<IQueryOption> getTopOption(int iSize) {
        ArrayList<IQueryOption> queryOptions = new ArrayList<IQueryOption>();
        queryOptions.add(new Top(iSize));
        return queryOptions;
    }
    // [End] 23474

    public static class List {

        // [Start] 24258: Support to load complete info only when required
        public static ArrayList<IQueryOption> getOption(int iMode) {
            /*
            if (MODE_DETAIL == iMode) {
                return getDetailOption();
            }

            ArrayList<IQueryOption> queryOptions = new ArrayList<IQueryOption>();

            Expand expand = new Expand("RootFolder");
            Select select = new Select(new String[]{
                    "BaseTemplate",
                    "BaseType",
                    "Id",
                    "Title",
                    "Hidden",
                    "LastItemModifiedDate",
                    "HasUniqueRoleAssignments",
                    "ItemCount",
                    "RootFolder/ServerRelativeURL",
                    "ParentWebUrl"
            });

            queryOptions.add(expand);
            queryOptions.add(select);
            return queryOptions;
            */
            return getDetailOption();
        }

        private static ArrayList<IQueryOption> getDetailOption() {
            ArrayList<IQueryOption> queryOptions = new ArrayList<IQueryOption>();

            Expand expand = new Expand("DataSource", "RootFolder", "RootFolder/Properties");
            Select select = new Select(new String[]{
                    "*",
                    "AllowDeletion",
                    "OnQuickLaunch",
                    "HasUniqueRoleAssignments",
                    "DataSource",
                    "RootFolder/ServerRelativeURL",
                    "RootFolder/Properties/TimelineDefaultView",
                    "RootFolder/Properties/TimelineAllViews",
                    "RootFolder/Properties/Timeline_Timeline"
            });
            // [Start] 23699: Add support to External Data
            expand.add("Fields");
            select.add(new String[]{
                    // only query for min required value
                    "Fields/InternalName", "Fields/FieldTypeKind", "Fields/TypeAsString", "Fields/SchemaXml"
            });
            // [End] 23699

            queryOptions.add(expand);
            queryOptions.add(select);
            return queryOptions;
        }
        // [End] 24258

        public static IFilterRestriction getHiddenRestriction(boolean bHidden) {
            return new IsEqualTo("Hidden", bHidden);
        }
    }

    // [Start] 23203: Added ListData query
    public static class ListData {

        public static ArrayList<IQueryOption> getOption(ArrayList<IFilterRestriction> alRestriction, String[] aSelect) {
            ArrayList<IQueryOption> queryOptions = new ArrayList<IQueryOption>();
            if (alRestriction.size() > 1) {
                queryOptions.add(new Filter(new And(alRestriction)));
            } else if (alRestriction.size() == 1) {
                queryOptions.add(new Filter(alRestriction.get(0)));
            }
            queryOptions.add(new Select(aSelect));
            return queryOptions;
        }

        public static ArrayList<IQueryOption> getOption(String sPath, String sContentType) {
            ArrayList<IFilterRestriction> alRestriction = new ArrayList<IFilterRestriction>();
            alRestriction.add(getPathRestriction(sPath));
            alRestriction.add(getContentTypeRestriction(sContentType));
            return getOption(alRestriction, new String[]{"Name", "Modified", "Created", "Path", "Id"});
        }

        public static ArrayList<IQueryOption> getIdOption(String sPath) {
            ArrayList<IFilterRestriction> alRestriction = new ArrayList<IFilterRestriction>();
            alRestriction.add(getPathRestriction(sPath));
            return getOption(alRestriction, new String[]{"Id"});
        }

        // 23474: Added id greater than restriction
        public static IFilterRestriction getIdGtRestriction(int id) {
            return new IsGreaterThan("Id", id);
        }

        public static IFilterRestriction getIdRestriction(int id) {
            return new IsEqualTo("Id", id);
        }

        public static IFilterRestriction getTitleRestriction(String sTitle) {
            return new IsEqualTo("Title", sTitle);
        }

        public static IFilterRestriction getPathRestriction(String sPath) {
            return new IsEqualTo("Path", sPath);
        }

        public static IFilterRestriction getNameRestriction(String sName) {
            return new IsEqualTo("Name", sName);
        }

        public static IFilterRestriction getContentTypeRestriction(String sContentType) {
            return new IsEqualTo("ContentType", sContentType);
        }
    }
    // [End] 23203

    public static class ListItem {

        // [Start] 24258: Support to load complete info only when required
        public static ArrayList<IQueryOption> getOption(com.independentsoft.share.List list, int iMode) {
            // return getDetailOption(list);
            // 25054: Temp list all on user information list
            if (MODE_DETAIL == iMode || list.isUserInformationList()) {
                return getDetailOption(list);
            }

            ArrayList<IQueryOption> queryOptions = new ArrayList<IQueryOption>();

            Expand expand = new Expand(new String[]{"FieldValuesAsText", "File", "File/Properties"});
            Select select = new Select(new String[]{"FileSystemObjectType", "HasUniqueRoleAssignments", "Id", "FileRef", "FileLeafRef"});
            if (list != null) {
                select.add(new String[]{"Modified", "HTML_x0020_File_x0020_Type"});
                // [Start] 23535: Support item versions
                if (list.isVersioningEnabled()) {
                    expand.add(new String[]{"File/Versions"});
                }
                // [End] 23535
                // 23597: AppRequest list does not support AttachmentFiles but it has attachments enabled
                // if (list != null && list.hasAttachmentsEnabled()) {
                if (list.hasAttachmentsEnabled() && !list.isAppRequestList()) {
                    expand.add("AttachmentFiles");
                }

                if (list.isDocumentLibraryBaseType()) {
                    select.add(new String[]{"File/Properties/vti_x005f_filesize"});
                    select.add(new String[]{"OData__UIVersionString"});
                } else {
                    select.add("Title");
                    if (!list.isSurvey()) {
                        select.add("GUID");
                    }
                }
                // [Start] 22494: Add support to Calendar list content
                if (list != null && list.isEventList()) {
                    select.add(new String[]{"RecurrenceData", "EventType", "TimeZone", "MasterSeriesItemID", "UID"});
                }
                // [End] 22494

                // [Start] 23699: Add support to External Data to get default missing field value
                if (list != null) {
                    for (com.independentsoft.share.List.FieldDefinition definition : list.getFieldDefinition()) {
                        if (com.independentsoft.share.Field.isExternalData(definition.getTypeAsString())) {
                            String sField = definition.getExtDataRelatedField();
                            if (sField != null && !"".equals(sField)) {
                                select.add(sField);
                            }
                        }
                    }
                }
                // [End] 23699
            }

            queryOptions.add(expand);
            queryOptions.add(select);
            return queryOptions;
        }

        private static ArrayList<IQueryOption> getDetailOption(com.independentsoft.share.List list) {
            ArrayList<IQueryOption> queryOptions = new ArrayList<IQueryOption>();

            Expand expand = new Expand(new String[]{"FieldValuesAsText", "File", "File/Properties"});
            // [Start] 23535: Support item versions
            if (list != null && list.isVersioningEnabled()) {
                expand.add(new String[]{"Versions", "File/Versions"});
            }
            // [End] 23535
            // 23597: AppRequest list does not support AttachmentFiles but it has attachments enabled
            // if (list != null && list.hasAttachmentsEnabled()) {
            if (list != null && list.hasAttachmentsEnabled() && !list.isAppRequestList()) {
                expand.add("AttachmentFiles");
            }

            // 23763: add HasUniqueRoleAssignments
            // Select select = new Select(new String[]{"*", "FileRef", "FileLeafRef"});
            Select select = new Select(new String[]{"*", "HasUniqueRoleAssignments", "FileRef", "FileLeafRef"});
            // [Start] 22494: Add support to Calendar list content
            if (list != null && list.isEventList()) {
                select.add(new String[]{"RecurrenceData", "EventType", "TimeZone", "MasterSeriesItemID", "UID"});
            }
            // [End] 22494

            // [Start] 23699: Add support to External Data to get default missing field value
            if (list != null) {
                for (com.independentsoft.share.List.FieldDefinition definition : list.getFieldDefinition()) {
                    if (com.independentsoft.share.Field.isExternalData(definition.getTypeAsString())) {
                        String sField = definition.getExtDataRelatedField();
                        if (sField != null && !"".equals(sField)) {
                            select.add(sField);
                        }
                    }
                }
            }
            // [End] 23699

            queryOptions.add(expand);
            queryOptions.add(select);
            return queryOptions;
        }
        // [End] 24258

        public static IFilterRestriction getFileDirRefRestriction(String sFileDirRef) {
            return new IsEqualTo("FileDirRef", sFileDirRef);
        }

        // [Start] 23322: Add leaf reference restriction
        public static IFilterRestriction getFileLeafRefRestriction(String sFileLeafRef) {
            return new IsEqualTo("FileLeafRef", sFileLeafRef);
        }
        // [End] 23322

        public static IFilterRestriction getFileRefRestriction(String sFileRef) {
            return new IsEqualTo("FileRef", sFileRef);
        }

        public static IFilterRestriction getIdRestriction(String sId) {
            return new IsEqualTo("Id", sId);
        }

        public static IFilterRestriction getIdRestriction(java.util.List<String> alId) {
            if (alId.size() == 1) {
                return new IsEqualTo("Id", alId.get(0));
            }
            ArrayList<IFilterRestriction> restrictionList = new ArrayList<IFilterRestriction>();
            for (String sId : alId) {
                restrictionList.add(new IsEqualTo("Id", sId));
            }
            return (new Or(restrictionList));
        }

        public static IFilterRestriction getIdRangeRestriction(long lGt, long lLt) {
            // lGt < Id < lLt
            return new And(new IsGreaterThan("Id", lGt), new IsLessThan("Id", lLt));
        }

        public static IFilterRestriction getIdGTRestriction(long lGt) {
            // lGt < Id < lLt
            return new IsGreaterThan("Id", lGt);
        }

        public static IFilterRestriction getGuidRestriction(String sGuid) {
            return new IsEqualTo("GUID", sGuid);
        }

        public static IFilterRestriction getTitleRestriction(String sTitle) {
            return new IsEqualTo("Title", sTitle);
        }

        // [Start] 21984: Add to get list item shareInfo option
        public static ArrayList<IQueryOption> getSharingInformationOption() {
            ArrayList<IQueryOption> queryOptions = new ArrayList<IQueryOption>();
            queryOptions.add(new Expand("permissionsInformation"));
            queryOptions.add(new Select("permissionsInformation"));
            return queryOptions;
        }
        // [End] 21984

        public static ArrayList<IQueryOption> getMaxIdOption() {
            ArrayList<IQueryOption> queryOptions = new ArrayList<IQueryOption>();

            queryOptions.add(new Top(1));
            queryOptions.add(new OrderBy(new PropertyOrder("Id", true)));
            return queryOptions;
        }

        public static ArrayList<IQueryOption> getMinIdOption() {
            ArrayList<IQueryOption> queryOptions = new ArrayList<IQueryOption>();

            queryOptions.add(new Top(1));
            queryOptions.add(new OrderBy(new PropertyOrder("Id", false)));
            return queryOptions;
        }

        public static ArrayList<IQueryOption> getPaginationOption(int iLastItemId, int iSize, ArrayList<IFilterRestriction> alRestriction) {
            if (iLastItemId >= 0) {
                alRestriction.add(new IsGreaterThan("Id", iLastItemId));
            }
            return getPaginationOption(iSize, alRestriction);
        }

        public static ArrayList<IQueryOption> getPaginationOption(int iSize, ArrayList<IFilterRestriction> alRestriction) {
            ArrayList<IQueryOption> queryOptions = new ArrayList<IQueryOption>();
            if (alRestriction.size() > 1) {
                queryOptions.add(new Filter(new And(alRestriction)));
            } else if (alRestriction.size() > 0) {
                queryOptions.add(new Filter(alRestriction.get(0)));
            }
            queryOptions.add(new Top(iSize));
            queryOptions.add(new OrderBy(new PropertyOrder("Id", false)));
            return queryOptions;
        }
    }

    public static class RoleAssignment {

        public static ArrayList<IQueryOption> getOption() {
            ArrayList<IQueryOption> queryOptions = new ArrayList<IQueryOption>();

            queryOptions.add(new Expand("Member", "RoleDefinitionBindings"));
            return queryOptions;
        }
    }

    public static class Site {

        // [Start] 24258: Support to load complete info only when required
        public static ArrayList<IQueryOption> getOption(int iMode) {
            if (MODE_DETAIL == iMode) {
                return getDetailOption();
            }

            ArrayList<IQueryOption> queryOptions = new ArrayList<IQueryOption>();

            queryOptions.add(new Select(new String[]{
                    "Id",
                    "Url",
                    "Created",
                    "ServerRelativeUrl",
                    "WebTemplate",
                    "Title",
                    // [Start] 23626: Additional fields for restore
                    "SiteLogoUrl",
                    "QuickLaunchEnabled",
                    "TreeViewEnabled"
                    // [End] 23626
            }));

            return queryOptions;
        }

        private static ArrayList<IQueryOption> getDetailOption() {
            // [Start] 22014: Modified to get timezone info
            // return new ArrayList<IQueryOption>();
            ArrayList<IQueryOption> queryOptions = new ArrayList<IQueryOption>();

            queryOptions.add(new Expand("RegionalSettings/TimeZone"));
            queryOptions.add(new Select("*", "RegionalSettings/TimeZone/Id"));

            return queryOptions;
            // [End] 22014
        }
        // [End] 24258

        public static IFilterRestriction getUrlRestriction(String sUrl) {
            // filter the url not pointing to the same host. This issue comes from Access app.
            // The url is https://ahsay-8b77212453c56d.sharepoint.com/proj_sub/access app
            // At the moment, it does not support back up apps data.
            return new StartsWith("Url", sUrl, false);
        }
    }

    // [Start] 22014: Added site properties query options
    public static class SiteProperties {

        public static ArrayList<IQueryOption> getOption() {
            ArrayList<IQueryOption> queryOptions = new ArrayList<IQueryOption>();
            queryOptions.add(new Expand("Owner", "usage"));
            // queryOptions.add(new Select("Owner/Email","usage/Storage","usage/StoragePercentageUsed"));
            return queryOptions;
        }

        public static IFilterRestriction getUrlRestriction(String sUrl) {
            return new StartsWith("Url", sUrl, false);
        }
    }
    // [End] 22014

    public static class SiteTemplate {

        public static ArrayList<IQueryOption> getOption() {
            return new ArrayList<IQueryOption>();
        }

        // public static IFilterRestriction getHiddenRestriction(boolean bHidden) {
        public static IFilterRestriction getHiddenRestriction() {
            return new IsEqualTo("isHidden", false);
        }
    }

    public static class Folder {

        public static ArrayList<IQueryOption> getOption() {
            return new ArrayList<IQueryOption>();
        }

        public static ArrayList<IQueryOption> getChildrenIdOption() {
            ArrayList<IQueryOption> queryOptions = new ArrayList<IQueryOption>();
            queryOptions.add(new Expand("Folders/ListItemAllFields", "Files/ListItemAllFields"));
            queryOptions.add(new Select("*", "Folders/ListItemAllFields/Id", "Files/ListItemAllFields/Id"));
            return queryOptions;
        }

        // 23474: Modified to get desired field only for listRestorer
        public static ArrayList<IQueryOption> getListItemInfoOption() {
            ArrayList<IQueryOption> queryOptions = new ArrayList<IQueryOption>();
            queryOptions.add(new Expand("ListItemAllFields"));
            queryOptions.add(new Select("*", "ListItemAllFields/Id", "ListItemAllFields/Title", "ListItemAllFields/FileRef"));
            return queryOptions;
        }

        /*
        // filter "Forms" folders as it is unnecessary at the moment
        IsNotEqualTo ne = new IsNotEqualTo("Name", "Forms");
        IsNotEqualTo ne2 = new IsNotEqualTo("Name", "_t"); // the thumbnails seem to be under a hidden folder in the Picture Library called _t, and the previews are under _w
        IsNotEqualTo ne3 = new IsNotEqualTo("Name", "_w");
        And and = new And(ne, ne2, ne3);
        Filter filter = new Filter(and);
        queryOptions.add(filter);
        */
    }

    public static class File {

        public static ArrayList<IQueryOption> getOption() {
            return new ArrayList<IQueryOption>();
        }

        public static ArrayList<IQueryOption> getIdOption() {
            ArrayList<IQueryOption> queryOptions = new ArrayList<IQueryOption>();
            queryOptions.add(new Expand("ListItemAllFields"));
            queryOptions.add(new Select("*", "ListItemAllFields/Id"));
            return queryOptions;
        }

        public static ArrayList<IQueryOption> getMinIdOption() {
            ArrayList<IQueryOption> queryOptions = new ArrayList<IQueryOption>();
            queryOptions.addAll(getIdOption());
            queryOptions.add(new Skip(100));
            return queryOptions;
        }
    }

    public static class Field {

        public static ArrayList<IQueryOption> getOption() {
            return new ArrayList<IQueryOption>();
        }

        public static ArrayList<IQueryOption> getIdOption() {
            ArrayList<IQueryOption> queryOptions = new ArrayList<IQueryOption>();
            queryOptions.add(new Filter(getIdRestriction()));
            return queryOptions;
        }

        public static IFilterRestriction getIdRestriction() {
            return new IsEqualTo("InternalName", "FirstNamePhonetic");
        }

        // [Start] 23322: Add internal name restriction
        public static IFilterRestriction getInternalNameRestriction(String sName) {
            return new IsEqualTo("InternalName", sName);
        }
        // [End] 23322

        public static IFilterRestriction getHiddenRestriction(boolean bHidden) {
            if (bHidden) {
                return new Or(
                        new IsEqualTo("Hidden", bHidden),
                        new IsEqualTo("Group", "_Hidden")
                );
            } else {
                return new And(
                        new IsEqualTo("Hidden", bHidden),
                        new IsNotEqualTo("Group", "_Hidden")
                );
            }
        }

        public static IFilterRestriction getUniqueValueRestriction() {
            return new IsEqualTo("EnforceUniqueValues", true);
        }
    }

    public static class ContentType {

        public static ArrayList<IQueryOption> getOption() {
            ArrayList<IQueryOption> queryOptions = new ArrayList<IQueryOption>();

            queryOptions.add(new Expand(new String[]{"Parent", "Fields", "WorkflowAssociations"}));
            queryOptions.add(new Select(new String[]{"*",
                    "Parent/Id", "Parent/Name", "Parent/Group", "Parent/Hidden",
                    // only query for min required value
                    "Fields/Id", "Fields/InternalName", "Fields/Group", "Fields/Hidden",
            }));
            return queryOptions;
        }

        // [Start] 23322: Add name restriction
        public static IFilterRestriction getNameRestriction(String sName) {
            return new IsEqualTo("Name", sName);
        }
        // [End] 23322

        public static IFilterRestriction getHiddenRestriction(boolean bHidden) {
            if (bHidden) {
                return new Or(
                        new IsEqualTo("Hidden", bHidden),
                        new IsEqualTo("Group", "_Hidden")
                );
            } else {
                return new And(
                        new IsEqualTo("Hidden", bHidden),
                        new IsNotEqualTo("Group", "_Hidden")
                );
            }
        }
    }

    // [Start] 22256: Get owner info
    public static class Group {

        public static ArrayList<IQueryOption> getOption() {
            ArrayList<IQueryOption> queryOptions = new ArrayList<IQueryOption>();

            // just expanding the owner will throw exception
            // if select loginName under owner, no exception thrown
            // queryOptions.add(new Expand("Users"));
            queryOptions.add(new Expand("Users", "Owner"));
            queryOptions.add(new Select("*", "Owner/LoginName"));
            return queryOptions;
        }
    }
    // [End] 22256

    public static class NavigationNode {

        public static ArrayList<IQueryOption> getOption() {
            ArrayList<IQueryOption> queryOptions = new ArrayList<IQueryOption>();

            queryOptions.add(new Expand(new String[]{"Children"}));
            queryOptions.add(new Select(new String[]{"*", "Children/Id"}));
            return queryOptions;
        }
    }

    // [Start] 24258: Support role options control
    public static class Role {

        public static ArrayList<IQueryOption> getOption() {
            return new ArrayList<IQueryOption>();
        }
    }
    // [End] 24258
}
