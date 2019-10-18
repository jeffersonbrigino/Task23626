package com.independentsoft.share;

import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;
import java.io.InputStream;
import java.text.ParseException;
import java.util.Date;

/*
 * Copyright (c) 2018 Ahsay Systems Corporation Limited. All Rights Reserved.
 *
 * Description: The Class SiteInfo.
 *
 * Date        Task  Author            Changes
 * 2019-10-18  23626 jefferson.brigino Added field siteLogoUrl
 */

public class SiteInfo {

    private int configuration;
    private Date createdTime;
    private String description;
    private String id;
    private Locale language = Locale.getInstance(Locale.Defines.NONE);
    private Date lastItemModifiedTime;
    private String serverRelativeUrl;
    private String title;
    private String webTemplate;
    private int webTemplateId;
    // 23626: Added siteLogoUrl for support on backup and restore
    private String siteLogoUrl;

    /**
     * Instantiates a new site info.
     */
    public SiteInfo() {
    }

    SiteInfo(InputStream inputStream) throws XMLStreamException, ParseException {
        XMLInputFactory xmlInputFactory = XMLInputFactory.newInstance();
        XMLStreamReader reader = xmlInputFactory.createXMLStreamReader(inputStream);

        parse(reader);
    }

    SiteInfo(XMLStreamReader reader) throws XMLStreamException, ParseException {
        parse(reader);
    }

    private void parse(XMLStreamReader reader) throws XMLStreamException, ParseException {
        while (reader.hasNext()) {
            if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("properties") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices/metadata")) {
                while (reader.hasNext()) {
                    if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("Configuration") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices")) {
                        String stringValue = reader.getElementText();

                        if (stringValue != null && stringValue.length() > 0) {
                            configuration = Integer.parseInt(stringValue);
                        }
                    } else if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("Created") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices")) {
                        String stringValue = reader.getElementText();

                        if (stringValue != null && stringValue.length() > 0) {
                            createdTime = Util.parseDate(stringValue);
                        }
                    } else if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("Description") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices")) {
                        description = reader.getElementText();
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
                    } else if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("ServerRelativeUrl") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices")) {
                        serverRelativeUrl = reader.getElementText();
                    } else if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("Title") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices")) {
                        title = reader.getElementText();
                    } else if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("WebTemplate") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices")) {
                        webTemplate = reader.getElementText();
                    } else if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("WebTemplateId") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices")) {
                        String stringValue = reader.getElementText();

                        if (stringValue != null && stringValue.length() > 0) {
                            webTemplateId = Integer.parseInt(stringValue);
                        }
                    // [Start] 23626: Added siteLogoUrl for support on backup and restore
                    } else if (reader.isStartElement() && reader.getLocalName() != null && reader.getNamespaceURI() != null && reader.getLocalName().equals("SiteLogoUrl") && reader.getNamespaceURI().equals("http://schemas.microsoft.com/ado/2007/08/dataservices")) {
                        siteLogoUrl = reader.getElementText();
                    }
                    // [End] 23626
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
     * Gets the description.
     *
     * @return the description
     */
    public String getDescription() {
        return description;
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
     * Gets the server relative url.
     *
     * @return the server relative url
     */
    public String getServerRelativeUrl() {
        return serverRelativeUrl;
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
     * Gets the web template.
     *
     * @return the web template
     */
    public String getWebTemplate() {
        return webTemplate;
    }

    /**
     * Gets the web template id.
     *
     * @return the web template id
     */
    public int getWebTemplateId() {
        return webTemplateId;
    }

    public String getSiteLogoUrl() {
        return siteLogoUrl;
    }

    public void setSiteLogoUrl(String siteLogoUrl) {
        this.siteLogoUrl = siteLogoUrl;
    }

    @Override
    public String toString() {
        return "SiteInfo{" +
                "configuration=" + configuration +
                ", createdTime=" + createdTime +
                ", description='" + description + '\'' +
                ", id='" + id + '\'' +
                ", language=" + language +
                ", lastItemModifiedTime=" + lastItemModifiedTime +
                ", serverRelativeUrl='" + serverRelativeUrl + '\'' +
                ", title='" + title + '\'' +
                ", webTemplate='" + webTemplate + '\'' +
                ", webTemplateId=" + webTemplateId +
                // 23626: Added siteLogoUrl for support on backup and restore
                ", SiteLogoUrl=" + siteLogoUrl +
                '}';
    }
}
