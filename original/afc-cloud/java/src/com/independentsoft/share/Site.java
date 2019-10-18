package com.independentsoft.share;

import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;
import java.io.IOException;
import java.text.ParseException;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.Map;

/*
 * Copyright (c) 2018 Ahsay Systems Corporation Limited. All Rights Reserved.
 *
 * Description: Site
 *
 * Date        Task  Author           Changes
 * 2018-03-09 20630  nicholas.leung   Created
 * 2019-03-27 23411  felix.chou       Support create site object from custom Xml
 * 2019-04-17 22014  felix.chou       Support storing more elements in site object
 * 2019-04-29 23625  felix.chou       Added methods to get features
 * 2019-05-06 23535  nicholas.leung   Support timezone
 * 2019-07-15 24258  nicholas.leung   Support site properities
 */
public class Site
        extends RawXmlEntity
        implements Comparable<Site> {

    public static final String ATTR_SITE_FEATURES = "AttrSFeats";
    public static final String ATTR_FEATURES = "AttrFeats";
    // 24258: Support site properities
    public static final String ATTR_SITE_PROPERTIES = "AttrSProps";
    // 23535: Support timezone
    public static final String ATTR_TIMEZONE = "AttrTz";

    private boolean allowRssFeeds;
    private String appInstanceId;
    private int configuration;
    private Date createdTime;
    private String customMasterUrl;
    private String description;
    private boolean isDocumentLibraryCalloutOfficeWebAppPreviewersDisabled;
    private boolean enableMinimalDownload;
    private String id;
    private Locale language = Locale.getInstance(Locale.Defines.NONE);
    private Date lastItemModifiedTime;
    private String masterUrl;
    private boolean isQuickLaunchEnabled;
    private boolean isRecycleBinEnabled;
    private String serverRelativeUrl;
    private boolean isSyndicationEnabled;
    private String title;
    private boolean isTreeViewEnabled;
    private int uiVersion;
    private boolean isUIVersionConfigurationEnabled;
    private boolean isOverwriteTranslationsOnChange;
    private String url;
    private String webTemplate;

    private Folder folder;
    TimeZone timeZone = null;

    /**
     * Instantiates a new site.
     */
    public Site() {
    }

    @Override
    public void setAttributeFromXml(AttributeType attrType, String sXml)
            throws XMLStreamException, ParseException, IOException {
        if (attrType.getType().equals(ATTR_SITE_FEATURES)
                || attrType.getType().equals(ATTR_FEATURES)
                // 24258: Support site properities
                || attrType.getType().equals(ATTR_SITE_PROPERTIES)
                // 23535: Support timezone
                || attrType.getType().equals(ATTR_TIMEZONE)) {
            return;
        }
        super.setAttributeFromXml(attrType, sXml);
    }

    @Override
    protected void parse(AttributeType attrType, XMLStreamReader reader)
            throws XMLStreamException, ParseException {
        while (reader.hasNext()) {
            if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("inline") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices/metadata")) {
                parseInline(reader);
            } else if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("properties") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices/metadata")) {
                while (reader.hasNext()) {
                    if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("AllowRssFeeds") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices")) {
                        String stringValue = reader.getElementText();

                        if (stringValue != null && stringValue.length() > 0) {
                            allowRssFeeds = Boolean.parseBoolean(stringValue);
                        }
                    } else if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("AppInstanceId") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices")) {
                        appInstanceId = reader.getElementText();
                    } else if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("Configuration") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices")) {
                        String stringValue = reader.getElementText();

                        if (stringValue != null && stringValue.length() > 0) {
                            configuration = Integer.parseInt(stringValue);
                        }
                    } else if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("Created") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices")) {
                        String stringValue = reader.getElementText();

                        if (stringValue != null && stringValue.length() > 0) {
                            createdTime = Util.parseDate(stringValue);
                        }
                    } else if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("CustomMasterUrl") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices")) {
                        customMasterUrl = reader.getElementText();
                    } else if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("Description") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices")) {
                        description = reader.getElementText();
                    } else if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("DocumentLibraryCalloutOfficeWebAppPreviewersDisabled") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices")) {
                        String stringValue = reader.getElementText();

                        if (stringValue != null && stringValue.length() > 0) {
                            isDocumentLibraryCalloutOfficeWebAppPreviewersDisabled = Boolean.parseBoolean(stringValue);
                        }
                    } else if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("EnableMinimalDownload") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices")) {
                        String stringValue = reader.getElementText();

                        if (stringValue != null && stringValue.length() > 0) {
                            enableMinimalDownload = Boolean.parseBoolean(stringValue);
                        }
                    } else if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("Id") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices")) {
                        id = reader.getElementText();
                    } else if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("Language") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices")) {
                        String stringValue = reader.getElementText();

                        if (stringValue != null && stringValue.length() > 0) {
                            language = Locale.getInstance(stringValue);
                        }
                    } else if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("LastItemModifiedDate") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices")) {
                        String stringValue = reader.getElementText();

                        if (stringValue != null && stringValue.length() > 0) {
                            lastItemModifiedTime = Util.parseDate(stringValue);
                        }
                    } else if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("MasterUrl") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices")) {
                        masterUrl = reader.getElementText();
                    } else if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("QuickLaunchEnabled") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices")) {
                        String stringValue = reader.getElementText();

                        if (stringValue != null && stringValue.length() > 0) {
                            isQuickLaunchEnabled = Boolean.parseBoolean(stringValue);
                        }
                    } else if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("RecycleBinEnabled") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices")) {
                        String stringValue = reader.getElementText();

                        if (stringValue != null && stringValue.length() > 0) {
                            isRecycleBinEnabled = Boolean.parseBoolean(stringValue);
                        }
                    } else if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("ServerRelativeUrl") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices")) {
                        serverRelativeUrl = reader.getElementText();
                    } else if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("SyndicationEnabled") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices")) {
                        String stringValue = reader.getElementText();

                        if (stringValue != null && stringValue.length() > 0) {
                            isSyndicationEnabled = Boolean.parseBoolean(stringValue);
                        }
                    } else if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("Title") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices")) {
                        title = reader.getElementText();
                    } else if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("TreeViewEnabled") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices")) {
                        String stringValue = reader.getElementText();

                        if (stringValue != null && stringValue.length() > 0) {
                            isTreeViewEnabled = Boolean.parseBoolean(stringValue);
                        }
                    } else if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("UIVersion") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices")) {
                        String stringValue = reader.getElementText();

                        if (stringValue != null && stringValue.length() > 0) {
                            uiVersion = Integer.parseInt(stringValue);
                        }
                    } else if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("UIVersionConfigurationEnabled") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices")) {
                        String stringValue = reader.getElementText();

                        if (stringValue != null && stringValue.length() > 0) {
                            isUIVersionConfigurationEnabled = Boolean.parseBoolean(stringValue);
                        }
                    } else if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("OverwriteTranslationsOnChange") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices")) {
                        String stringValue = reader.getElementText();

                        if (stringValue != null && stringValue.length() > 0) {
                            isOverwriteTranslationsOnChange = Boolean.parseBoolean(stringValue);
                        }
                    } else if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("Url") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices")) {
                        url = reader.getElementText();
                    } else if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("WebTemplate") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices")) {
                        webTemplate = reader.getElementText();
                    }

                    if (reader.isEndElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("properties") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices/metadata")) {
                        break;
                    } else {
                        reader.next();
                    }
                }
            }

            if (reader.isEndElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("entry") && reader.getNamespaceURI().equals("http://www.w3.org/2005/Atom")) {
                break;
            } else {
                reader.next();
            }
        }
    }

    // [Start] 22014: Parse more information
    private void parseInline(XMLStreamReader reader)
            throws XMLStreamException, ParseException {
        while (reader.hasNext()) {
            if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("entry") && reader.getNamespaceURI().equals("http://www.w3.org/2005/Atom")) {
                while (reader.hasNext()) {
                    if (reader.isStartElement() && reader.getLocalName() != null && reader.getLocalName().equals("id")) {
                        String id = reader.getElementText();
                        if (id.endsWith("/RegionalSettings")) {
                            parseRegionalSettings(reader);
                        }
                    }
                    if (reader.isEndElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("entry") && reader.getNamespaceURI().equals("http://www.w3.org/2005/Atom")) {
                        break;
                    } else {
                        reader.next();
                    }
                }
            }
            if (reader.isEndElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("inline") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices/metadata")) {
                break;
            } else {
                reader.next();
            }
        }
    }

    private void parseRegionalSettings(XMLStreamReader reader)
            throws XMLStreamException, ParseException {
        while (reader.hasNext()) {
            if (reader.isStartElement() && reader.getLocalName() != null && reader.getLocalName().equals("id")) {
                String id = reader.getElementText();
                if (id.endsWith("/RegionalSettings/TimeZone")) {
                    timeZone = new TimeZone();
                    timeZone.parse(reader);
                }
            }
            if (reader.isEndElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("inline") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices/metadata")) {
                break;
            } else {
                reader.next();
            }
        }
    }
    // [End] 22014

    public String toJSon(Map<String, FieldValue> mFieldValue) {
        return new FieldValue.ObjectValue("SP.Web", mFieldValue).toString();
    }

    public String toUpdateJSon() {
        Map<String, FieldValue> mValue = new LinkedHashMap<String, FieldValue>();
        if (customMasterUrl != null && customMasterUrl.length() > 0) {
            mValue.put("CustomMasterUrl", new FieldValue.StringValue(customMasterUrl));
        }
        if (description != null && description.length() > 0) {
            mValue.put("Description", new FieldValue.StringValue(description));
        }
        mValue.put("EnableMinimalDownload", new FieldValue.StringValue(enableMinimalDownload));
        if (masterUrl != null && masterUrl.length() > 0) {
            mValue.put("MasterUrl", new FieldValue.StringValue(masterUrl));
        }
        mValue.put("QuickLaunchEnabled", new FieldValue.StringValue(isQuickLaunchEnabled));
        if (serverRelativeUrl != null && serverRelativeUrl.length() > 0) {
            mValue.put("ServerRelativeUrl", new FieldValue.StringValue(serverRelativeUrl));
        }
        mValue.put("SyndicationEnabled", new FieldValue.StringValue(isSyndicationEnabled));
        if (title != null && title.length() > 0) {
            mValue.put("Title", new FieldValue.StringValue(title));
        }
        mValue.put("TreeViewEnabled", new FieldValue.StringValue(isTreeViewEnabled));
        if (uiVersion > 0) {
            mValue.put("UIVersion", new FieldValue.StringValue(uiVersion));
        }
        mValue.put("UIVersionConfigurationEnabled", new FieldValue.StringValue(isUIVersionConfigurationEnabled));
        return toJSon(mValue);
    }

    /**
     * Checks if is rss feeds allowed.
     *
     * @return true, if is rss feeds allowed
     */
    public boolean isRssFeedsAllowed() {
        return allowRssFeeds;
    }

    /**
     * Gets the app instance id.
     *
     * @return the app instance id
     */
    public String getAppInstanceId() {
        return appInstanceId;
    }

    /**
     * Gets the configuration.
     *
     * @return the configuration
     */
    public int getConfiguration() {
        return configuration;
    }

    /**
     * Gets the created time.
     *
     * @return the created time
     */
    public Date getCreatedTime() {
        return createdTime;
    }

    /**
     * Gets the custom master url.
     *
     * @return the custom master url
     */
    public String getCustomMasterUrl() {
        return customMasterUrl;
    }

    /**
     * Sets the custom master url.
     *
     * @param customMasterUrl the new custom master url
     */
    public void setCustomMasterUrl(String customMasterUrl) {
        this.customMasterUrl = customMasterUrl;
    }

    /**
     * Gets the description.
     *
     * @return the description
     */
    public String getDescription() {
        return description;
    }

    /**
     * Sets the description.
     *
     * @param description the new description
     */
    public void setDescription(String description) {
        this.description = description;
    }

    /**
     * Checks if is document library callout office web app previewers disabled.
     *
     * @return true, if is document library callout office web app previewers disabled
     */
    public boolean isDocumentLibraryCalloutOfficeWebAppPreviewersDisabled() {
        return isDocumentLibraryCalloutOfficeWebAppPreviewersDisabled;
    }

    /**
     * Checks if is minimal download enabled.
     *
     * @return true, if is minimal download enabled
     */
    public boolean isMinimalDownloadEnabled() {
        return enableMinimalDownload;
    }

    /**
     * Enable minimal download.
     *
     * @param enableMinimalDownload the enable minimal download
     */
    public void enableMinimalDownload(boolean enableMinimalDownload) {
        this.enableMinimalDownload = enableMinimalDownload;
    }

    /**
     * Gets the id.
     *
     * @return the id
     */
    public String getId() {
        return id;
    }

    /**
     * Gets the language.
     *
     * @return the language
     */
    public Locale getLanguage() {
        return language;
    }

    /**
     * Gets the last item modified time.
     *
     * @return the last item modified time
     */
    public Date getLastItemModifiedTime() {
        return lastItemModifiedTime;
    }

    /**
     * Gets the master url.
     *
     * @return the master url
     */
    public String getMasterUrl() {
        return masterUrl;
    }

    /**
     * Sets the master url.
     *
     * @param masterUrl the new master url
     */
    public void setMasterUrl(String masterUrl) {
        this.masterUrl = masterUrl;
    }

    /**
     * Checks if is quick launch enabled.
     *
     * @return true, if is quick launch enabled
     */
    public boolean isQuickLaunchEnabled() {
        return isQuickLaunchEnabled;
    }

    /**
     * Enable quick launch.
     *
     * @param isQuickLaunchEnabled the is quick launch enabled
     */
    public void enableQuickLaunch(boolean isQuickLaunchEnabled) {
        this.isQuickLaunchEnabled = isQuickLaunchEnabled;
    }

    /**
     * Checks if is recycle bin enabled.
     *
     * @return true, if is recycle bin enabled
     */
    public boolean isRecycleBinEnabled() {
        return isRecycleBinEnabled;
    }

    /**
     * Gets the server relative url.
     *
     * @return the server relative url
     */
    public String getServerRelativeUrl() {
        return serverRelativeUrl;
    }

    /**
     * Sets the server relative url.
     *
     * @param serverRelativeUrl the new server relative url
     */
    public void setServerRelativeUrl(String serverRelativeUrl) {
        this.serverRelativeUrl = serverRelativeUrl;
    }

    /**
     * Checks if is syndication enabled.
     *
     * @return true, if is syndication enabled
     */
    public boolean isSyndicationEnabled() {
        return isSyndicationEnabled;
    }

    /**
     * Enable syndication.
     *
     * @param isSyndicationEnabled the is syndication enabled
     */
    public void enableSyndication(boolean isSyndicationEnabled) {
        this.isSyndicationEnabled = isSyndicationEnabled;
    }

    /**
     * Gets the title.
     *
     * @return the title
     */
    public String getTitle() {
        return title;
    }

    /**
     * Sets the title.
     *
     * @param title the new title
     */
    public void setTitle(String title) {
        this.title = title;
    }

    /**
     * Checks if is tree view enabled.
     *
     * @return true, if is tree view enabled
     */
    public boolean isTreeViewEnabled() {
        return isTreeViewEnabled;
    }

    /**
     * Enable tree view.
     *
     * @param isTreeViewEnabled the is tree view enabled
     */
    public void enableTreeView(boolean isTreeViewEnabled) {
        this.isTreeViewEnabled = isTreeViewEnabled;
    }

    /**
     * Gets the UI version.
     *
     * @return the UI version
     */
    public int getUIVersion() {
        return uiVersion;
    }

    /**
     * Sets the UI version.
     *
     * @param uiVersion the new UI version
     */
    public void setUIVersion(int uiVersion) {
        this.uiVersion = uiVersion;
    }

    /**
     * Checks if is UI version configuration enabled.
     *
     * @return true, if is UI version configuration enabled
     */
    public boolean isUIVersionConfigurationEnabled() {
        return isUIVersionConfigurationEnabled;
    }

    public boolean isOverwriteTranslationsOnChange() {
        return isOverwriteTranslationsOnChange;
    }

    /**
     * Enable ui version configuration.
     *
     * @param isUIVersionConfigurationEnabled the is ui version configuration enabled
     */
    public void enableUIVersionConfiguration(boolean isUIVersionConfigurationEnabled) {
        this.isUIVersionConfigurationEnabled = isUIVersionConfigurationEnabled;
    }

    // [Start] 23625: get features
    public ManagedFeatures getWebFeatures()
            throws Exception {
        return getFeatures(ATTR_FEATURES);
    }

    public ManagedFeatures getSiteFeatures()
            throws Exception {
        return getFeatures(ATTR_SITE_FEATURES);
    }

    private ManagedFeatures getFeatures(String sAttrType)
            throws Exception {
        String sXml = getRawDescriptor().getString(sAttrType);
        if (sXml == null) {
            return null;
        }
        SerializableJSONObject jo = new SerializableJSONObject(getRawDescriptor().getString(sAttrType));
        ManagedFeatures managedFeatures = new ManagedFeatures();
        managedFeatures.setAttribute(jo.getString(ATTR_DEFAULT));
        return managedFeatures;
    }
    // [End] 23625

    // [Start] 24258: Support site properities
    public SiteProperties getSiteProperties() {
        try {
            SiteProperties siteProperties = new SiteProperties();
            siteProperties.parseFromDescriptor(getRawDescriptor().getString(ATTR_SITE_PROPERTIES));
            return siteProperties;
        } catch (Throwable t) {
        }
        return null;
    }
    // [End] 24258

    // [Start] 23535: Support timezone
    public TimeZone getTimeZone() {
        try {
            TimeZone tz = new TimeZone();
            tz.parseFromDescriptor(getRawDescriptor().getString(ATTR_TIMEZONE));
            return tz;
        } catch (Throwable t) {
        }
        return timeZone;
    }
    // [End] 23535

    /**
     * Gets the url.
     *
     * @return the url
     */
    public String getUrl() {
        return url;
    }

    public void setUrl(String url) {
        this.url = url;
    }

    /**
     * Gets the web template.
     *
     * @return the web template
     */
    public String getWebTemplate() {
        return webTemplate;
    }

    public void setWebTemplate(String webTemplate) {
        this.webTemplate = webTemplate;
    }

    public Folder getFolder() {
        return folder;
    }

    public void setFolder(Folder folder) {
        this.folder = folder;
    }

    // [Start] 23411: support create site object from custom xml
    @Override
    public int compareTo(Site o) {
        return this.getUrl().compareTo(o.getUrl());
    }

    public void setCustomXml()
            throws Exception {
        setAttribute("<?xml version=\"1.0\" encoding=\"utf-8\"?>\n" +
                "<entry xml:base=\"\" xmlns=\"http://www.w3.org/2005/Atom\" xmlns:d=\"http://schemas.microsoft.com/ado/2007/08/dataservices\" xmlns:m=\"http://schemas.microsoft.com/ado/2007/08/dataservices/metadata\" xmlns:georss=\"http://www.georss.org/georss\" xmlns:gml=\"http://www.opengis.net/gml\">\n" +
                "\t<id />\n" +
                "\t<category term=\"SP.Web\" scheme=\"http://schemas.microsoft.com/ado/2007/08/dataservices/scheme\" />\n" +
                "\t<link rel=\"edit\" href=\"Web\" />\n" +
                "\t<title />\n" +
                "\t<updated />\n" +
                "\t<author>\n" +
                "\t\t<name />\n" +
                "\t</author>\n" +
                "\t<content type=\"application/xml\">\n" +
                "\t\t<m:properties>\n" +
                "\t\t\t<d:Url>" + url + "</d:Url>\n" +
                "\t\t\t<d:WebTemplate>" + webTemplate + "</d:WebTemplate>\n" +
                "\t\t</m:properties>\n" +
                "\t</content>\n" +
                "</entry>"
        );
    }
    // [End] 23411

    @Override
    public String toString() {
        return "Site{" +
                "allowRssFeeds=" + allowRssFeeds +
                ", appInstanceId='" + appInstanceId + '\'' +
                ", configuration=" + configuration +
                ", createdTime=" + createdTime +
                ", customMasterUrl='" + customMasterUrl + '\'' +
                ", description='" + description + '\'' +
                ", isDocumentLibraryCalloutOfficeWebAppPreviewersDisabled=" + isDocumentLibraryCalloutOfficeWebAppPreviewersDisabled +
                ", enableMinimalDownload=" + enableMinimalDownload +
                ", id='" + id + '\'' +
                ", language=" + language +
                ", lastItemModifiedTime=" + lastItemModifiedTime +
                ", masterUrl='" + masterUrl + '\'' +
                ", isQuickLaunchEnabled=" + isQuickLaunchEnabled +
                ", isRecycleBinEnabled=" + isRecycleBinEnabled +
                ", serverRelativeUrl='" + serverRelativeUrl + '\'' +
                ", isSyndicationEnabled=" + isSyndicationEnabled +
                ", title='" + title + '\'' +
                ", isTreeViewEnabled=" + isTreeViewEnabled +
                ", uiVersion=" + uiVersion +
                ", isUIVersionConfigurationEnabled=" + isUIVersionConfigurationEnabled +
                ", isOverwriteTranslationsOnChange=" + isOverwriteTranslationsOnChange +
                ", url='" + url + '\'' +
                ", webTemplate='" + webTemplate + '\'' +
                ", folder='" + folder + '\'' +
                '}';
    }
}