package com.ahsay.afc.cloud.office365.sharepoint;

import com.independentsoft.share.*;
import com.independentsoft.share.Locale;
import com.independentsoft.share.queryoptions.IQueryOption;

import java.util.*;
import java.util.List;

/*
 * Copyright (c) 2018 Ahsay Systems Corporation Limited. All Rights Reserved.
 *
 * Description: SharePoint Unit Test
 *
 * Date        Task  Author           Changes
 */
public class SharePointUnitTest {

    public static void main(String[] args) {
        try {
            Service.Callback callback = new Service.Callback() {
                @Override
                public boolean isDebug() {
                    return true;
                }

                @Override
                public int getMsSleepTimeForNextRetry(int iRetryCount, Exception fault) {
                    if (fault instanceof ServiceException) {
                        ServiceException svrFault = (ServiceException) fault;
                        System.out.println("Code: " + svrFault.getStatusCode());
                        int iRetryAfter = svrFault.getRetryAfter();
                        if (iRetryAfter >= 0) {
                            System.out.println("Retry after: " + iRetryAfter);
                            return iRetryAfter * 1000;
                        }
                        if (svrFault.hasErrorResponse()) {
                            // no need to retry if get valid response from server
                            return -1;
                        }
                        int iStatusCode = svrFault.getStatusCode();
                        if (iStatusCode == 404) {
                            return -1;
                        }
                    }
                    if (iRetryCount > 4) {
                        return -1;
                    }
                    // default retry time
                    System.out.println("Retry after: " + (iRetryCount * 20 / 1000));
                    return iRetryCount * 20;
                }
            };

            String sUser = "carven.tsang@cloudbacko.biz";
            String sPwd = "Cde123$%";

            String[] aURL = {
                    // "https://ahsay.sharepoint.com"
                    // "https://ahsay.sharepoint.com/sites/UnitTest"
                     "https://ahsay.sharepoint.com/sites/RestoreTests"
//                    "https://ahsay.sharepoint.com/sites/customer_intelligence"
                    // "https://ahsay.sharepoint.com/sites/UnitTest/Collaboration_TermSite",
                    // "https://ahsay-my.sharepoint.com/personal/carven_tsang_cloudbacko_biz",
//                    "https://ahsay.sharepoint.com/AllLists",
            };

            for (String sURL : aURL) {
                Service service = new Service(sURL, sUser, sPwd, null, callback);
                String sSiteURL = service.getSiteUrl();

                Site site = service.getSite(sSiteURL, SharePointQueryUtils.Site.getOption(SharePointQueryUtils.MODE_OVERVIEW));
                System.out.println("--->> " + sSiteURL + " <<--- (" + site.getTitle() + ") - " + site.getId());
                System.out.println("INIT: " + site.toString());

//                site.setTitle("Updated Title");
//                service.updateSite(sSiteURL, site);
//
//                Site site2 = service.getSite(sSiteURL, SharePointQueryUtils.Site.getOption(SharePointQueryUtils.MODE_OVERVIEW));
//                System.out.println("--->> " + sSiteURL + " <<--- (" + site2.getTitle() + ") - " + site2.getId());
//                System.out.println("UPDATED: " + site2.toString());

//                CurrentThemeInfo currentThemeInfo = service.getCurrentThemeInfo(sSiteURL);
//                System.out.println(currentThemeInfo.getFontSchemeUrl());
//                System.out.println(currentThemeInfo.getBackgroundImageUrl());
//                System.out.println(currentThemeInfo.getColorPaletteUrl());

//                String schemeUrl = "/sites/qatest2/_catalogs/theme/Themed/B1E4A0E2/theme.spfont";
//                String colorPalletteUrl = "/sites/qatest2/_catalogs/theme/Themed/B1E4A0E2/theme.spcolor";
//                String backgroundImageUrl = null;
//
//                System.out.println("APPLY THEME: " + service.applyTheme(sSiteURL, colorPalletteUrl));
//                List<NavigationNode> quickLaunch = service.getQuickLaunch(sSiteURL);
//
//                for (NavigationNode n: quickLaunch) {
//                    System.out.println(n.toString());
//                }
//
//                NavigationNode navigationNode = new NavigationNode(true, false, true, "SampleTitle", "/sites/RestoreTests/_layouts/15/viewlsts.aspx");
//                service.addQuickLaunch(sSiteURL, navigationNode);
//                List<NavigationNode> topNavigationBar = service.getTopNavigationBar(sSiteURL);
//                for (NavigationNode n: topNavigationBar) {
//                    System.out.println(n.toString());
//                }

//                List<NavigationNode> nodeChildren = service.getNavigationNodeChildren(sSiteURL, 1033);
//                for (NavigationNode n: nodeChildren) {
//                    System.out.println(n.toString());
//                }
//                getManagedFeatures(service, sSiteURL);
//                System.out.println("");
//                getAvailableFields(service, sSiteURL);
//                System.out.println("");
//                getAvailableContentTypes(service, sSiteURL);
//                System.out.println("");
//                getLists(service, sSiteURL);
            }
        } catch (Throwable t) {
            t.printStackTrace();
        }
    }

    private static void getManagedFeatures(Service service, String sSiteURL)
            throws Exception {
        System.out.println("------------ Site ManagedFeatures ------------");
        ArrayList<ManagedFeatures.Feature> features = FeatureUtils.Site.getManagedFeatures(service, sSiteURL).getFeatures();
        for (ManagedFeatures.Feature feature : features) {
            System.out.println("[" + (feature.isActivated() ? "X" : " ") + "] " + feature.getDisplayName() + " - " + feature.getId());
        }
        System.out.println("------------ Web ManagedFeatures ------------");
        features = FeatureUtils.Web.getManagedFeatures(service, sSiteURL).getFeatures();
        for (ManagedFeatures.Feature feature : features) {
            System.out.println("[" + (feature.isActivated() ? "X" : " ") + "] " + feature.getDisplayName() + " - " + feature.getId());
        }
    }

    private static void getAvailableFields(Service service, String sSiteURL)
            throws Exception {
        System.out.println("------------ AvailableFields ------------");
        int iOffset = 0;
        int iMax = Constant.DEFAULT_RETURN_ITEM;
        while (true) {
            ArrayList<IQueryOption> queryOptions = SharePointQueryUtils.Field.getOption();
            queryOptions.addAll(SharePointQueryUtils.getPaginationOption(iOffset, iMax));
            List<Field> fields = service.getAvailableFields(sSiteURL, queryOptions);
            if (fields == null) {
                break;
            }
            Collections.sort(fields, new Comparator<Field>() {
                @Override
                public int compare(Field f1, Field f2) {
                    if (f1 == null) {
                        return f2 == null ? 0 : -1;
                    } else if (f2 == null) {
                        return 1;
                    }

                    String s1 = f1.getGroup() + "-" + f1.getTitle();
                    String s2 = f2.getGroup() + "-" + f2.getTitle();
                    return s1.compareTo(s2);
                }
            });
            for (int i = 0; i < fields.size(); i++) {
                Field field = fields.get(i);
                System.out.println(i + ": [" + (Field.isBuiltIn(field) ? " " : "X") + "]"
                        + " [" + (field.isSealed() ? "S" : "N") + "]"
                        + " " + field.getTitle() + " (" + field.getGroup() + ") - " + field.getInternalName()
                        + " [" + (field.isHidden() ? "H" : "V") + "]"
                        + " [" + (field.isReadOnly() ? "R" : "W") + "]"
                        + " [" + (field.getSchemaXml().isFromBaseType() ? "B" : "?") + "]"
                        + " @ " + field.getScope() + " - " + field.getSchemaXml().getSourceId()
                );
            }
            if (fields.size() < iMax) {
                break;
            }
            iOffset += fields.size();
        }
    }

    private static void getAvailableContentTypes(Service service, String sSiteURL)
            throws Exception {
        System.out.println("------------ AvailableContentTypes ------------");
        int iOffset = 0;
        int iMax = Constant.DEFAULT_RETURN_ITEM;
        while (true) {
            ArrayList<IQueryOption> queryOptions = SharePointQueryUtils.ContentType.getOption();
            queryOptions.addAll(SharePointQueryUtils.getPaginationOption(iOffset, iMax));
            List<ContentType> contentTypes = service.getAvailableContentTypes(sSiteURL, queryOptions);
            // List<ContentType> contentTypes = service.getContentTypes(sSiteURL, queryOptions);
            if (contentTypes == null) {
                break;
            }
            Collections.sort(contentTypes, new Comparator<ContentType>() {
                @Override
                public int compare(ContentType c1, ContentType c2) {
                    if (c1 == null) {
                        return c2 == null ? 0 : -1;
                    } else if (c2 == null) {
                        return 1;
                    }

                    String s1 = c1.getGroup() + "-" + c1.getName();
                    String s2 = c2.getGroup() + "-" + c2.getName();
                    return s1.compareTo(s2);
                }
            });
            for (int i = 0; i < contentTypes.size(); i++) {
                ContentType contentType = contentTypes.get(i);
                System.out.println(i + ": [" + (ContentType.isBuiltIn(contentType) ? " " : "X") + "] "
                        + " [" + (contentType.isSealed() ? "S" : "N") + "]"
                        + " " + contentType.getName() + " (" + contentType.getGroup() + ")"
                        + " [" + (contentType.isHidden() ? "H" : "V") + "]"
                        + " [" + (contentType.isReadOnly() ? "R" : "W") + "]"
                        + " @ " + contentType.getScope()
                );
            }
            if (contentTypes.size() < iMax) {
                break;
            }
            iOffset += contentTypes.size();
        }
    }

    private static void getLists(Service service, String sSiteURL)
            throws Exception {
        System.out.println("------------ Lists ------------");
        List<com.independentsoft.share.List> lists = service.getLists(sSiteURL, SharePointQueryUtils.List.getOption(SharePointQueryUtils.MODE_OVERVIEW));
        Collections.sort(lists, new Comparator<com.independentsoft.share.List>() {
            @Override
            public int compare(com.independentsoft.share.List l1, com.independentsoft.share.List l2) {
                if (l1 == null) {
                    return l2 == null ? 0 : -1;
                } else if (l2 == null) {
                    return 1;
                }

                String s1 = l1.getBaseType().getName() + "-" + l1.getTitle();
                String s2 = l2.getBaseType().getName() + "-" + l2.getTitle();
                return s1.compareTo(s2);
            }
        });
        ArrayList<String> alSettings = new ArrayList<String>();
        for (int i = 0; i < lists.size(); i++) {
            com.independentsoft.share.List list = lists.get(i);
            System.out.println(i + ": [" + (list.isHidden() ? " " : "X") + "] "
                    + "[" + (list.isAllowDeletion() ? "W" : "R") + "] "
                    + list.getTitle() + " (" + list.getBaseType().getName() + ")" + " (" + list.getTemplateType().getName() + ")"
                    + " " + list.getServerRelativeUrl()
            );
            ListSettings settings = new ListSettings();
            for (ListSettings.Type type : ListSettings.Type.values()) {
                Map<String, ListSettings.GetLink> mLink = service.getListSettingsAvailableLinks(
                        sSiteURL, list.getId(), list.hasUniqueRoleAssignments(), type, settings
                );
                if (mLink != null) {
                    for (ListSettings.GetLink link : mLink.values()) {
                        if (link instanceof ListSettings.GetLink.Html) {
                            if (!alSettings.contains(((ListSettings.GetLink.Html) link).getName())) {
                                alSettings.add(((ListSettings.GetLink.Html) link).getName());
                            }
                        }
                    }
                }
            }
        }
        System.out.println("------------ AvailableListSettings ------------");
        for (String s : alSettings) {
            System.out.println(s);
        }
    }
}