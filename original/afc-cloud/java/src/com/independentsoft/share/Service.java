package com.independentsoft.share;

import com.ahsay.afc.cloud.office365.sharepoint.Constant;
import com.independentsoft.share.queryoptions.IQueryOption;

import javax.xml.stream.XMLStreamException;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.SocketTimeoutException;
import java.text.ParseException;
import java.util.*;
import java.util.List;

/*
 * Copyright (c) 2018 Ahsay Systems Corporation Limited. All Rights Reserved.
 *
 * Description: Service class contains the methods to perform operations on SharePoint server.
 *
 * Date        Task  Author           Changes
 * 2018-03-09 20630  nicholas.leung   Remove dupicate logic
 * 2018-10-09 22245  nicholas.leung   Update REST call to support % and # in files and folders
 * 2018-10-09 22234  nicholas.leung   Modify site collections query option
 * 2018-10-31 22363  terry.li         Added method to getSite without retry
 * 2018-12-04 22245  felix.chou       Handle ' in filePath
 * 2018-12-12 22692  felix.chou       Use decodedUrl in method create folder/file
 * 2019-01-09 21984  felix.chou       Add methods to get and set item sharing info
 * 2019-02-21 23203  felix.chou       Added method to getListData
 * 2019-03-20 23322  felix.chou       Added option to get site user info list
 * 2019-03-27 23411  felix.chou       Modified to get Site Collection template
 * 2019-04-03 23474  felix.chou       Support query options when create file/folder
 * 2019-04-18 22014  felix.chou       Added to get site properties
 * 2019-05-06 23535  nicholas.leung   Support another way to update document field value
 * 2015-05-31 22215  nicholas.leung   Support list item version control
 * 2019-06-06 22869  tszkin.seto      Added to get list item checkout user
 * 2019-07-08 23940  pong.tse         Changed to handle activateSiteFeature exceptions in this class
 *                                    and added method to check whether a site feature is activated
 * 2019-07-15 24258  nicholas.leung   Support to get single role assignment by id
 * 2019-08-19 24468  gavin.fu         For createTeamSite, use "Lcid" instead of "SPSiteLanguage"
 */
public class Service
        extends ServiceInstance {

    public static class Callback
            extends ServiceInstance.Callback {

        @Override
        public String getPrintDebugName() {
            return "Service";
        }
    }

    public Service(String siteUrl, String username, String password, String region) {
        super(siteUrl, username, password, region, new Callback());
    }

    public Service(String siteUrl, String username, String password, String region, Callback callback) {
        super(siteUrl, username, password, region, callback);
    }

    public Service(String siteUrl, String username, String password, String domain, String region) {
        super(siteUrl, username, password, domain, region, new Callback());
    }

    public Service(String siteUrl, String username, String password, String domain, String region, Callback callback) {
        super(siteUrl, username, password, domain, region, callback);
    }

    @Override
    public Callback getCallback() {
        return (Callback) callback;
    }

    /**
     * Gets the folders.
     *
     * @param parentFolder the parent folder
     * @return the folders
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<Folder> getFolders(String siteUrl, String parentFolder) throws ServiceException {
        return getFolders(siteUrl, parentFolder, null);
    }

    /**
     * Gets the folders.
     *
     * @param parentFolder the parent folder
     * @param queryOptions the query options
     * @return the folders
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<Folder> getFolders(String siteUrl, String parentFolder, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (parentFolder == null) {
            throw new IllegalArgumentException("parentFolder");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        // queryOptionsString = queryOptionsString.isEmpty() ? queryOptionsString : "&" + queryOptionsString.substring(1);
        // 22245: Support % and # in files and folders
        // StringBuilder requestUrl = new StringBuilder("_api/web/GetFolderByServerRelativeUrl(@v)/folders?@v='" + Util.encodeUrl(parentFolder) + "'" + queryOptionsString;
        // StringBuilder requestUrl = new StringBuilder("_api/web/GetFolderByServerRelativePath(decodedUrl=@v)/folders?@v='" + Util.encodeUrl(Util.escapeQueryUrl(parentFolder)) + "'" + queryOptionsString;
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFolderByServerRelativePath(decodedUrl=@v)/folders?@v='" + Util.encodeUrl(Util.escapeQueryUrl(parentFolder)) + "'");
        if (sbQuery.length() > 0) {
            requestUrl.append("&");
            requestUrl.append(sbQuery.substring(1));
        }
        if (callback.isDebug()) {
            callback.printDebug("getFolders", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.FoldersHandler handler = new ServiceResponseUtil.FoldersHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getFolders();
    }

    /**
     * Gets the files.
     *
     * @param folderPath the folder path
     * @return the files
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<File> getFiles(String siteUrl, String folderPath) throws ServiceException {
        return getFiles(siteUrl, folderPath, null);
    }

    /**
     * Gets the files.
     *
     * @param folderPath   the folder path
     * @param queryOptions the query options
     * @return the files
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<File> getFiles(String siteUrl, String folderPath, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (folderPath == null) {
            throw new IllegalArgumentException("folderPath");
        }

        /*
        Util.queryOptionsToString(sbQuery, queryOptions);
        queryOptionsString = queryOptionsString.isEmpty() ? queryOptionsString : "&" + queryOptionsString.substring(1);
        // 22245: Support % and # in files and folders
        // StringBuilder requestUrl = new StringBuilder("_api/web/GetFolderByServerRelativeUrl(@v)/files?@v='" + Util.encodeUrl(folderPath) + "'" + queryOptionsString;
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFolderByServerRelativePath(decodedUrl=@v)/files?@v='" + Util.encodeUrl(Util.escapeQueryUrl(folderPath)) + "'" + queryOptionsString;
        */
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFolderByServerRelativePath(decodedUrl=@v)/files?@v='" + Util.encodeUrl(Util.escapeQueryUrl(folderPath)) + "'");
        if (sbQuery.length() > 0) {
            requestUrl.append("&");
            requestUrl.append(sbQuery.substring(1));
        }
        if (callback.isDebug()) {
            callback.printDebug("getFiles", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.FilesHandler handler = new ServiceResponseUtil.FilesHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getFiles();
    }

    /**
     * Delete folder.
     *
     * @param folderPath the folder path
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public void deleteFolder(String siteUrl, String folderPath)
            throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (folderPath == null) {
            throw new IllegalArgumentException("folderPath");
        }

        // 22245: Support % and # in files and folders
        // StringBuilder requestUrl = new StringBuilder("_api/web/GetFolderByServerRelativeUrl(@v)?@v='" + Util.encodeUrl(folderPath) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFolderByServerRelativePath(decodedUrl=@v)?@v='" + Util.encodeUrl(Util.escapeQueryUrl(folderPath)) + "'");
        if (callback.isDebug()) {
            callback.printDebug("deleteFolder", siteUrl, requestUrl.toString());
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), null, "DELETE", null);
    }

    /**
     * Creates the folder.
     *
     * @param folderPath the folder path
     * @return the folder
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    // [Start] 23474: Support query option
    public Folder createFolder(String siteUrl, String folderPath)
            throws ServiceException {
        return createFolder(siteUrl, folderPath, null);
    }

    public Folder createFolder(String siteUrl, String folderPath, List<IQueryOption> queryOptions)
            throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (folderPath == null) {
            throw new IllegalArgumentException("folderPath");
        }

        // 22692: Use decodedUrl as folder path to handle special chars
        // StringBuilder requestUrl = new StringBuilder("_api/web/folders/add('" + Util.encodeUrl(folderPath) + "')");
        // StringBuilder requestUrl = new StringBuilder("_api/web/folders/AddUsingPath(decodedurl=@v)?@v='" + Util.encodeUrl(Util.escapeQueryUrl(folderPath)) + "'");
        /*
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/folders/AddUsingPath(decodedurl=@v)?@v='" + Util.encodeUrl(Util.escapeQueryUrl(folderPath)) + "'" + queryOptionsString;
        */
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/folders/AddUsingPath(decodedurl=@v)?@v='" + Util.encodeUrl(Util.escapeQueryUrl(folderPath)) + "'");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("createFolder", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.FolderHandler handler = new ServiceResponseUtil.FolderHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.getFolder();
    }
    // [End] 23474

    /**
     * Gets the list.
     *
     * @param listId the list id
     * @return the list
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    private com.independentsoft.share.List getList(String siteUrl, String listId)
            throws ServiceException {
        return getList(siteUrl, listId, null);
    }

    public com.independentsoft.share.List getList(String siteUrl, String listId, List<IQueryOption> queryOptions)
            throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }

        /*
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')" + queryOptionsString;
        */
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getList", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ListHandler handler = new ServiceResponseUtil.ListHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getList();
    }

    /**
     * Gets the list by title.
     *
     * @param title the title
     * @return the list by title
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    private com.independentsoft.share.List getListByTitle(String siteUrl, String title)
            throws ServiceException {
        return getListByTitle(siteUrl, title, null);
    }

    public com.independentsoft.share.List getListByTitle(String siteUrl, String title, List<IQueryOption> queryOptions)
            throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (title == null) {
            throw new IllegalArgumentException("title");
        }

        /*
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists/GetByTitle('" + Util.encodeUrl(title) + "')" + queryOptionsString;
        */
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists/GetByTitle('" + Util.encodeUrl(title) + "')");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getListByTitle", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ListHandler handler = new ServiceResponseUtil.ListHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getList();
    }

    public com.independentsoft.share.List getListByUrl(String siteUrl, String url, List<IQueryOption> queryOptions)
            throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (url == null) {
            throw new IllegalArgumentException("url");
        }

        /*
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/GetList('" + Util.encodeUrl(url) + "')" + queryOptionsString;
        */
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/GetList('" + Util.encodeUrl(url) + "')");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getListByUrl", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ListHandler handler = new ServiceResponseUtil.ListHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getList();
    }

    // [Start] 23203: Added method to getListData
    public List<ListData> getListData(String siteUrl, String folderPath, List<IQueryOption> queryOptions)
            throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (folderPath == null) {
            throw new IllegalArgumentException("folderPath");
        }

        // StringBuilder requestUrl = "_vti_bin/ListData.svc/" + Util.encodeUrl(Util.escapeQueryUrl(folderPath)) + Util.queryOptionsToString(queryOptions);
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_vti_bin/ListData.svc/" + Util.encodeUrl(Util.escapeQueryUrl(folderPath)));
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("listData", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ListDatasHandler handler = new ServiceResponseUtil.ListDatasHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getListData();
    }
    // [End] 23203

    /**
     * Gets the list items.
     *
     * @param listId the list id
     * @return the list items
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    private List<ListItem> getListItems(String siteUrl, String listId) throws ServiceException {
        return getListItems(siteUrl, listId, null);
    }

    /**
     * Gets the list items.
     *
     * @param listId       the list id
     * @param queryOptions the query options
     * @return the list items
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<ListItem> getListItems(String siteUrl, String listId, List<IQueryOption> queryOptions)
            throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }

        /*
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/items" + queryOptionsString;
        */
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/items");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getListItems", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ListItemsHandler handler = new ServiceResponseUtil.ListItemsHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getListItems();
    }

    public List<ListItem> getListItems(String siteUrl, String listId, String camlQuery, List<IQueryOption> queryOptions)
            throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }

        // StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/GetItems" + Util.queryOptionsToString(queryOptions);
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/GetItems");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getListItems", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ListItemsHandler handler = new ServiceResponseUtil.ListItemsHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), camlQuery, handler);
        return handler.getListItems();
    }

    /**
     * Suggest.
     *
     * @param query the query
     * @return the suggest result
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public SuggestResult suggest(String siteUrl, String query) throws ServiceException {
        return suggest(siteUrl, new SearchQuerySuggestion(query));
    }

    /**
     * Suggest.
     *
     * @param suggestion the suggestion
     * @return the suggest result
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public SuggestResult suggest(String siteUrl, SearchQuerySuggestion suggestion) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (suggestion == null) {
            throw new IllegalArgumentException("suggestion");
        }

        // StringBuilder requestUrl = "/_api/search/suggest?" + suggestion.toString();
        StringBuilder requestUrl = new StringBuilder("/_api/search/suggest?" + suggestion.toString());
        if (callback.isDebug()) {
            callback.printDebug("suggest", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.SuggestResultHandler handler = new ServiceResponseUtil.SuggestResultHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getSuggestResult();
    }

    /**
     * Search.
     *
     * @param restriction the restriction
     * @return the search result
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public SearchResult search(String siteUrl, com.independentsoft.share.fql.IRestriction restriction) throws ServiceException {
        return search(siteUrl, new SearchQuery(restriction));
    }

    /**
     * Search.
     *
     * @param restriction the restriction
     * @return the search result
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public SearchResult search(String siteUrl, com.independentsoft.share.kql.IRestriction restriction) throws ServiceException {
        return search(siteUrl, new SearchQuery(restriction));
    }

    /**
     * Search.
     *
     * @param query the query
     * @return the search result
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public SearchResult search(String siteUrl, String query) throws ServiceException {
        return search(siteUrl, new SearchQuery(query));
    }

    /**
     * Search.
     *
     * @param query the query
     * @return the search result
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public SearchResult search(String siteUrl, SearchQuery query) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (query == null) {
            throw new IllegalArgumentException("query");
        }

        StringBuilder requestUrl = new StringBuilder("/_api/search/postquery");
        String requestBody = query.toString();
        if (callback.isDebug()) {
            callback.printDebug("search", siteUrl, requestUrl.toString(), requestBody);
        }
        ServiceResponseUtil.SearchResultHandler handler = new ServiceResponseUtil.SearchResultHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, handler);
        return handler.getSearchResult();
    }

    /**
     * Gets the list item.
     *
     * @param listId the list id
     * @param itemId the item id
     * @return the list item
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public ListItem getListItem(String siteUrl, String listId, int itemId) throws ServiceException {
        return getListItem(siteUrl, listId, itemId, null);
    }

    public ListItem getListItem(String siteUrl, String listId, int itemId, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (itemId < 0) {
            throw new IllegalArgumentException("The parameter itemId must be non-negative.");
        }

        /*
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/items(" + itemId + ")" + queryOptionsString;
        */
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/items(" + itemId + ")");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getListItem", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ListItemHandler handler = new ServiceResponseUtil.ListItemHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getListItem();
    }

    /**
     * Gets the list item property as xml string.
     *
     * @param listId       the list id
     * @param itemId       the item id
     * @param propertyName the property name
     * @return the list item
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public String getListItemProperty(String siteUrl, String listId, int itemId, String propertyName) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (propertyName == null) {
            throw new IllegalArgumentException("propertyName");
        }
        if (itemId < 0) {
            throw new IllegalArgumentException("The parameter itemId must be non-negative.");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/items(" + itemId + ")/" + Util.encodeUrl(propertyName));
        if (callback.isDebug()) {
            callback.printDebug("getListItemProperty", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.StringHandler handler = new ServiceResponseUtil.StringHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getString();
    }
    /*
    **
     * Gets the field values.
     *
     * @param listId the list id
     * @param itemId the item id
     * @return the field values
     * @throws com.independentsoft.share.ServiceException the service exception
     *
    public List<FieldValue> getFieldValuesAsHtml(String siteUrl, String listId, int itemId) throws ServiceException {
        return getFieldValuesAsHtml(siteUrl, listId, itemId, null);
    }

    **
     * Gets the field values.
     *
     * @param listId       the list id
     * @param itemId       the item id
     * @param queryOptions the query options
     * @return the field values
     * @throws com.independentsoft.share.ServiceException the service exception
     *
    public List<FieldValue> getFieldValuesAsHtml(String siteUrl, String listId, int itemId, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (itemId < 0) {
            throw new IllegalArgumentException("The parameter itemId must be non-negative.");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/items(" + itemId + ")/FieldValuesAsHTML" + queryOptionsString;
        if (callback.isDebug()) {
            callback.printDebug("getFieldValuesAsHtml", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ListFieldValuesHandler handler = new ServiceResponseUtil.ListFieldValuesHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getListFieldValues();
    }

    **
     * Gets the field values.
     *
     * @param listId the list id
     * @param itemId the item id
     * @return the field values
     * @throws com.independentsoft.share.ServiceException the service exception
     *
    public List<FieldValue> getFieldValues(String siteUrl, String listId, int itemId) throws ServiceException {
        return getFieldValues(siteUrl, listId, itemId, null);
    }

    **
     * Gets the field values.
     *
     * @param listId       the list id
     * @param itemId       the item id
     * @param queryOptions the query options
     * @return the field values
     * @throws com.independentsoft.share.ServiceException the service exception
     *
    public List<FieldValue> getFieldValues(String siteUrl, String listId, int itemId, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (itemId < 0) {
            throw new IllegalArgumentException("The parameter itemId must be non-negative.");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/items(" + itemId + ")/FieldValuesAsText" + queryOptionsString;
        if (callback.isDebug()) {
            callback.printDebug("getFieldValues", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ListFieldValuesHandler handler = new ServiceResponseUtil.ListFieldValuesHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getListFieldValues();
    }
    */

    /**
     * Gets the list item attachments.
     *
     * @param listId the list id
     * @param itemId the item id
     * @return the list item attachments
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<Attachment> getListItemAttachments(String siteUrl, String listId, int itemId) throws ServiceException {
        return getListItemAttachments(siteUrl, listId, itemId, null);
    }

    /**
     * Gets the list item attachments.
     *
     * @param listId       the list id
     * @param itemId       the item id
     * @param queryOptions the query options
     * @return the list item attachments
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<Attachment> getListItemAttachments(String siteUrl, String listId, int itemId, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (itemId < 0) {
            throw new IllegalArgumentException("The parameter itemId must be non-negative.");
        }
        /*
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/items(" + itemId + ")/AttachmentFiles" + queryOptionsString;
        */
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/items(" + itemId + ")/AttachmentFiles");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getListItemAttachments", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ListAttachmentsHandler handler = new ServiceResponseUtil.ListAttachmentsHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getListAttachments();
    }

    /**
     * Creates the list item attachment.
     *
     * @param listId
     * @param itemId
     * @param fileName
     * @param buffer
     * @return
     * @throws com.independentsoft.share.ServiceException
     */
    public Attachment createListItemAttachment(String siteUrl, String listId, int itemId, String fileName, byte[] buffer) throws ServiceException {
        ByteArrayInputStream stream = new ByteArrayInputStream(buffer);
        return createListItemAttachment(siteUrl, listId, itemId, fileName, stream);
    }

    /**
     * Creates the list item attachment.
     *
     * @param listId
     * @param itemId
     * @param fileName
     * @param stream
     * @return
     * @throws com.independentsoft.share.ServiceException
     */
    public Attachment createListItemAttachment(String siteUrl, String listId, int itemId, String fileName, InputStream stream) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (itemId < 0) {
            throw new IllegalArgumentException("The parameter itemId must be non-negative.");
        }
        if (fileName == null) {
            throw new IllegalArgumentException("filePath");
        }
        if (stream == null) {
            throw new IllegalArgumentException("stream");
        }

        // 22245: Handle ' in fileName
        // StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/items(" + itemId + ")/AttachmentFiles/Add(filename='" + Util.encodeUrl(fileName) + "')");
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/items(" + itemId + ")/AttachmentFiles/Add(filename='" + Util.encodeUrl(Util.escapeQueryUrl(fileName)) + "')");
        if (callback.isDebug()) {
            callback.printDebug("createListItemAttachment", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ListAttachmentHandler handler = new ServiceResponseUtil.ListAttachmentHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), null, null, null, stream, true, false, handler);
        return handler.getListAttachment();
    }

    /**
     * Updates the list item attachment.
     *
     * @param listId
     * @param itemId
     * @param fileName
     * @param buffer
     * @throws com.independentsoft.share.ServiceException
     */
    public void updateListItemAttachment(String siteUrl, String listId, int itemId, String fileName, byte[] buffer) throws ServiceException {
        ByteArrayInputStream stream = new ByteArrayInputStream(buffer);
        updateListItemAttachment(siteUrl, listId, itemId, fileName, stream);
    }

    /**
     * Updates the list item attachment.
     *
     * @param listId
     * @param itemId
     * @param fileName
     * @param stream
     * @throws com.independentsoft.share.ServiceException
     */
    public void updateListItemAttachment(String siteUrl, String listId, int itemId, String fileName, InputStream stream) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (itemId < 0) {
            throw new IllegalArgumentException("The parameter itemId must be non-negative.");
        }
        if (fileName == null) {
            throw new IllegalArgumentException("filePath");
        }
        if (stream == null) {
            throw new IllegalArgumentException("stream");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/items(" + itemId + ")/AttachmentFiles(filename='" + Util.encodeUrl(fileName) + "')/$value");
        if (callback.isDebug()) {
            callback.printDebug("updateListItemAttachment", siteUrl, requestUrl.toString());
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), null, "PUT", null, stream, true, false, null);
    }

    /**
     * Gets the list item content type.
     *
     * @param listId the list id
     * @param itemId the item id
     * @return the list item content type
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    private ContentType getListItemContentType(String siteUrl, String listId, int itemId) throws ServiceException {
        return getListItemContentType(siteUrl, listId, itemId, null);
    }

    public ContentType getListItemContentType(String siteUrl, String listId, int itemId, List<IQueryOption> queryOptions) throws ServiceException {
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (itemId < 0) {
            throw new IllegalArgumentException("The parameter itemId must be non-negative.");
        }

        /*
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/items(" + itemId + ")/ContentType" + queryOptionsString;
        */
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/items(" + itemId + ")/ContentType");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getListItemContentType", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ContentTypeHandler handler = new ServiceResponseUtil.ContentTypeHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getContentType();
    }

    /**
     * Gets the user effective permissions.
     *
     * @param listId    the list id
     * @param loginName the login name
     * @return the user effective permissions
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<BasePermission.Defines> getUserEffectivePermissions(String siteUrl, String listId, String loginName) throws ServiceException {
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        if (loginName == null) {
            throw new IllegalArgumentException("loginName");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/getusereffectivepermissions(@v)?@v='" + Util.encodeUrl(loginName) + "'");
        if (callback.isDebug()) {
            callback.printDebug("getUserEffectivePermissions", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BasePermissionsHandler handler = new ServiceResponseUtil.BasePermissionsHandler("GetUserEffectivePermissions");
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getBasePermissions();
    }

    /**
     * Gets the user effective permissions.
     *
     * @param listId    the list id
     * @param itemId    the item id
     * @param loginName the login name
     * @return the user effective permissions
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<BasePermission.Defines> getUserEffectivePermissions(String siteUrl, String listId, int itemId, String loginName) throws ServiceException {
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (itemId < 0) {
            throw new IllegalArgumentException("The parameter itemId must be non-negative.");
        }
        if (loginName == null) {
            throw new IllegalArgumentException("loginName");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/items(" + itemId + ")/getusereffectivepermissions(@v)?@v='" + Util.encodeUrl(loginName) + "'");
        if (callback.isDebug()) {
            callback.printDebug("getUserEffectivePermissions", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BasePermissionsHandler handler = new ServiceResponseUtil.BasePermissionsHandler("GetUserEffectivePermissions");
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getBasePermissions();
    }

    /**
     * Delete list item.
     *
     * @param listId the list id
     * @param itemId the item id
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public void deleteListItem(String siteUrl, String listId, int itemId) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (itemId < 0) {
            throw new IllegalArgumentException("The parameter itemId must be non-negative.");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/items(" + itemId + ")");
        if (callback.isDebug()) {
            callback.printDebug("deleteListItem", siteUrl, requestUrl.toString());
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), null, "DELETE", "*", null);
    }

    // [Start] 23411: modified to get the webTemplate and return site list
    public List<Site> getSiteCollections(String siteUrl) throws Exception {
        ArrayList<Site> siteCollections = new ArrayList<Site>();
        int iOffset = 0;
        int iSize = 500;
        while (true) {
            List<Site> scs = getSiteCollections(siteUrl, iOffset, iSize);
            if (scs == null) {
                break;
            }
            siteCollections.addAll(scs);
            if (scs.size() < iSize) {
                break;
            }
            iOffset += iSize;
        }
        return siteCollections;
    }

    private List<Site> getSiteCollections(String siteUrl, int offset, int size) throws Exception {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        // ref: https://docs.microsoft.com/en-us/SharePoint/technical-reference/crawled-and-managed-properties-overview
        // String requestUrl = "/_api/search/query?querytext='contentclass:STS_Site+AND+-SPSiteURL:personal'&trimduplicates=false&startrow=" + offset + "&rowlimit=" + size;
        // String requestUrl = "/_api/search/query?querytext='contentclass:STS_Site+AND+-SPSiteURL:personal'&selectproperties='Path'&trimduplicates=true&startrow=" + offset + "&rowlimit=" + size;
        // String requestUrl = "/_api/search/query?querytext='contentclass:STS_Site+AND+-SPSiteURL:personal+AND+-SPSiteURL:portals'&selectproperties='Path'&trimduplicates=true&startrow=" + offset + "&rowlimit=" + size;
        // 22234: Use "trimduplicates=false" to get all results
        // StringBuilder requestUrl = "/_api/search/query?querytext='contentclass:STS_Site'&selectproperties='Path'&trimduplicates=true&startrow=" + offset + "&rowlimit=" + size;
        StringBuilder requestUrl = new StringBuilder("/_api/search/query?querytext='contentclass:STS_Site'&selectproperties='Path,webTemplate'&trimduplicates=false&startrow=" + offset + "&rowlimit=" + size);
        if (callback.isDebug()) {
            callback.printDebug("getSiteCollections", siteUrl, requestUrl.toString());
        }

        ArrayList<Site> siteCollections = new ArrayList<Site>();
        ServiceResponseUtil.SearchResultHandler handler = new ServiceResponseUtil.SearchResultHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        SearchResult searchResult = handler.getSearchResult();
        if (searchResult != null) {
            SimpleDataTable table = searchResult.getPrimaryQueryResult().getRelevantResult().getTable();
            for (SimpleDataRow row : table.getRows()) {
                Site siteCollection = new Site();
                for (KeyValue cell : row.getCells()) {
                    if (cell.getKey().equals("Path")) {
                        siteCollection.setUrl(cell.getValue());
                    }
                    if (cell.getKey().equals("webTemplate")) {
                        siteCollection.setWebTemplate(cell.getValue());
                    }
                }
                siteCollection.setCustomXml();
                siteCollections.add(siteCollection);
            }
        }
        return siteCollections;
    }
    // [End] 23411

    /**
     * Gets the site.
     *
     * @return the site
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    // [Start] 22363: Added option to send request with retry
    private Site getSite(String siteUrl)
            throws ServiceException {
        return getSite(siteUrl, null);
    }

    public Site getSite(String siteUrl, List<IQueryOption> queryOptions)
            throws ServiceException {
        return getSite(siteUrl, true, queryOptions);
    }

    // [Start] 22014: Added to support query
    public Site getSite(String siteUrl, boolean bRetryable, List<IQueryOption> queryOptions)
            throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        // StringBuilder requestUrl = new StringBuilder("_api/web" + Util.queryOptionsToString(queryOptions);
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getSite", siteUrl, requestUrl.toString());
        }
        // Added retry option to handle
        ServiceResponseUtil.SiteHandler handler = new ServiceResponseUtil.SiteHandler(bRetryable);
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getSite();
    }
    // [End] 22014
    // [End] 22363

    /**
     * Gets the sites.
     *
     * @return the sites
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    private List<Site> getSites(String siteUrl) throws ServiceException {
        return getSites(siteUrl, null);
    }

    /**
     * Gets the sites.
     *
     * @param queryOptions the query options
     * @return the sites
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<Site> getSites(String siteUrl, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        /*
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/Webs" + queryOptionsString;
        */
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/Webs");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getSites", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.SitesHandler handler = new ServiceResponseUtil.SitesHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getSites();
    }

    /**
     * Gets the site infos.
     *
     * @return the site infos
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<SiteInfo> getSiteInfos(String siteUrl) throws ServiceException {
        return getSiteInfos(siteUrl, null);
    }

    /**
     * Gets the site infos.
     *
     * @param queryOptions the query options
     * @return the site infos
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<SiteInfo> getSiteInfos(String siteUrl, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        /*
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/WebInfos" + queryOptionsString;
        */
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/WebInfos");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getSiteInfos", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.SiteInfosHandler handler = new ServiceResponseUtil.SiteInfosHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getSiteInfos();
    }

    // [Start] 22014: Method to get site owner and usage info
    public SiteProperties getSiteProperties(String sSiteUrl, List<IQueryOption> queryOptions)
            throws Exception {
        if (sSiteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        // StringBuilder requestUrl = "/_api/site" + Util.queryOptionsToString(queryOptions);
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/site");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getSiteProperties", sSiteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.SitePropertiesHandler handler = new ServiceResponseUtil.SitePropertiesHandler();
        doSendRequest(sSiteUrl, "GET", requestUrl.toString(), handler);
        return handler.getSiteProperties();
    }
    // [End] 22014

    /**
     * Apply theme.
     *
     * @param colorPaletteUrl the color palette url
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean applyTheme(String siteUrl, String colorPaletteUrl) throws ServiceException {
        return applyTheme(siteUrl, colorPaletteUrl, null);
    }

    /**
     * Apply theme.
     *
     * @param colorPaletteUrl the color palette url
     * @param shareGenerated  the share generated
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean applyTheme(String siteUrl, String colorPaletteUrl, boolean shareGenerated) throws ServiceException {
        return applyTheme(siteUrl, colorPaletteUrl, null, shareGenerated);
    }

    /**
     * Apply theme.
     *
     * @param colorPaletteUrl the color palette url
     * @param fontSchemeUrl   the font scheme url
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean applyTheme(String siteUrl, String colorPaletteUrl, String fontSchemeUrl) throws ServiceException {
        return applyTheme(siteUrl, colorPaletteUrl, fontSchemeUrl, null);
    }

    /**
     * Apply theme.
     *
     * @param colorPaletteUrl the color palette url
     * @param fontSchemeUrl   the font scheme url
     * @param shareGenerated  the share generated
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean applyTheme(String siteUrl, String colorPaletteUrl, String fontSchemeUrl, boolean shareGenerated) throws ServiceException {
        return applyTheme(siteUrl, colorPaletteUrl, fontSchemeUrl, null, shareGenerated);
    }

    /**
     * Apply theme.
     *
     * @param colorPaletteUrl    the color palette url
     * @param fontSchemeUrl      the font scheme url
     * @param backgroundImageUrl the background image url
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean applyTheme(String siteUrl, String colorPaletteUrl, String fontSchemeUrl, String backgroundImageUrl) throws ServiceException {
        return applyTheme(siteUrl, colorPaletteUrl, fontSchemeUrl, backgroundImageUrl, false);
    }

    /**
     * Apply theme.
     *
     * @param colorPaletteUrl    the color palette url
     * @param fontSchemeUrl      the font scheme url
     * @param backgroundImageUrl the background image url
     * @param shareGenerated     the share generated
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean applyTheme(String siteUrl, String colorPaletteUrl, String fontSchemeUrl, String backgroundImageUrl, boolean shareGenerated) throws ServiceException {
        if (colorPaletteUrl == null) {
            throw new IllegalArgumentException("colorPaletteUrl");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        String fontSchemeUrlParameter = "";
        String backgroundImageUrlParameter = "";
        if (fontSchemeUrl != null && fontSchemeUrl.length() > 0) {
            fontSchemeUrlParameter = ", 'fontSchemeUrl':'" + fontSchemeUrl + "'";
        }
        if (backgroundImageUrl != null && backgroundImageUrl.length() > 0) {
            backgroundImageUrlParameter = ", 'backgroundImageUrl':'" + backgroundImageUrl + "'";
        }
        StringBuilder requestUrl = new StringBuilder("_api/web/applytheme");
        StringBuilder requestBody = new StringBuilder("{'colorPaletteUrl':'" + colorPaletteUrl + "'" + fontSchemeUrlParameter + backgroundImageUrlParameter + ",'shareGenerated':" + Boolean.toString(shareGenerated).toLowerCase() + "}");
        // requestUrl = Util.encodeUrl(requestUrl);
        if (callback.isDebug()) {
            callback.printDebug("applyTheme", siteUrl, Util.encodeUrl(requestUrl.toString()), requestBody.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("ApplyTheme");
        doSendRequest(siteUrl, "POST", Util.encodeUrl(requestUrl.toString()), requestBody.toString(), handler);
        return handler.isSuccess();
    }

    /**
     * Apply site template.
     *
     * @param name the name
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean applySiteTemplate(String siteUrl, String name) throws ServiceException {
        if (name == null) {
            throw new IllegalArgumentException("name");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/ApplyWebTemplate(" + Util.encodeUrl(name) + ")");
        if (callback.isDebug()) {
            callback.printDebug("applySiteTemplate", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("ApplyWebTemplate");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.isSuccess();
    }

    /**
     * Break role inheritance.
     *
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean breakRoleInheritance(String siteUrl) throws ServiceException {
        return breakRoleInheritance(siteUrl, false);
    }

    /**
     * Break role inheritance.
     *
     * @param copyRoleAssignments the copy role assignments
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean breakRoleInheritance(String siteUrl, boolean copyRoleAssignments) throws ServiceException {
        return breakRoleInheritance(siteUrl, copyRoleAssignments, false);
    }

    /**
     * Break role inheritance.
     *
     * @param copyRoleAssignments the copy role assignments
     * @param clearSubscopes      the clear subscopes
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean breakRoleInheritance(String siteUrl, boolean copyRoleAssignments, boolean clearSubscopes) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/BreakRoleInheritance(copyroleassignments=" + Boolean.toString(copyRoleAssignments).toLowerCase() + ", clearsubscopes=" + Boolean.toString(clearSubscopes).toLowerCase() + ")");
        if (callback.isDebug()) {
            callback.printDebug("breakRoleInheritance", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("BreakRoleInheritance");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.isSuccess();
    }

    /**
     * Break role inheritance.
     *
     * @param listId the list id
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean breakRoleInheritance(String siteUrl, String listId) throws ServiceException {
        return breakRoleInheritance(listId, false);
    }

    /**
     * Break role inheritance.
     *
     * @param listId              the list id
     * @param copyRoleAssignments the copy role assignments
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean breakRoleInheritance(String siteUrl, String listId, boolean copyRoleAssignments) throws ServiceException {
        return breakRoleInheritance(listId, copyRoleAssignments, false);
    }

    /**
     * Break role inheritance.
     *
     * @param listId              the list id
     * @param copyRoleAssignments the copy role assignments
     * @param clearSubscopes      the clear subscopes
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean breakRoleInheritance(String siteUrl, String listId, boolean copyRoleAssignments, boolean clearSubscopes) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/BreakRoleInheritance(copyroleassignments=" + Boolean.toString(copyRoleAssignments).toLowerCase() + ", clearsubscopes=" + Boolean.toString(clearSubscopes).toLowerCase() + ")");
        if (callback.isDebug()) {
            callback.printDebug("breakRoleInheritance", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("BreakRoleInheritance");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.isSuccess();
    }

    /**
     * Break role inheritance.
     *
     * @param listId the list id
     * @param itemId the item id
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean breakRoleInheritance(String siteUrl, String listId, int itemId) throws ServiceException {
        return breakRoleInheritance(siteUrl, listId, itemId, false);
    }

    /**
     * Break role inheritance.
     *
     * @param listId              the list id
     * @param itemId              the item id
     * @param copyRoleAssignments the copy role assignments
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean breakRoleInheritance(String siteUrl, String listId, int itemId, boolean copyRoleAssignments) throws ServiceException {
        return breakRoleInheritance(siteUrl, listId, itemId, copyRoleAssignments, false);
    }

    /**
     * Break role inheritance.
     *
     * @param listId              the list id
     * @param itemId              the item id
     * @param copyRoleAssignments the copy role assignments
     * @param clearSubscopes      the clear subscopes
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean breakRoleInheritance(String siteUrl, String listId, int itemId, boolean copyRoleAssignments, boolean clearSubscopes) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/items(" + itemId + ")/BreakRoleInheritance(copyroleassignments=" + Boolean.toString(copyRoleAssignments).toLowerCase() + ", clearsubscopes=" + Boolean.toString(clearSubscopes).toLowerCase() + ")");
        if (callback.isDebug()) {
            callback.printDebug("breakRoleInheritance", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("BreakRoleInheritance");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.isSuccess();
    }

    /**
     * Ensure user.
     *
     * @param loginName the login name
     * @return the user
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public User ensureUser(String siteUrl, String loginName) throws ServiceException {
        if (loginName == null) {
            throw new IllegalArgumentException("loginName");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/EnsureUser");
        String requestBody = "{ 'logonName': '" + Util.encodeJson(loginName) + "' }";
        if (callback.isDebug()) {
            callback.printDebug("ensureUser", siteUrl, requestUrl.toString(), requestBody);
        }
        ServiceResponseUtil.UserHandler handler = new ServiceResponseUtil.UserHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, handler);
        return handler.getUser();
    }

    /**
     * Gets the catalog.
     *
     * @param listTemplate the list template
     * @return the catalog
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public com.independentsoft.share.List getCatalog(String siteUrl, ListTemplateType listTemplate) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/GetCatalog(" + listTemplate.getValue() + ")");
        if (callback.isDebug()) {
            callback.printDebug("getCatalog", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ListHandler handler = new ServiceResponseUtil.ListHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.getList();
    }

    /**
     * Gets the changes.
     *
     * @param query the query
     * @return the changes
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<Change> getChanges(String siteUrl, ChangeQuery query) throws ServiceException {
        return getChanges(siteUrl, query, new ArrayList<IQueryOption>());
    }

    /**
     * Gets the changes.
     *
     * @param query        the query
     * @param queryOptions the query options
     * @return the changes
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<Change> getChanges(String siteUrl, ChangeQuery query, List<IQueryOption> queryOptions) throws ServiceException {
        if (query == null) {
            throw new IllegalArgumentException("query");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        /*
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/GetChanges" + queryOptionsString;
        */
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/GetChanges");
        requestUrl.append(sbQuery);
        String requestBody = query.toString();
        if (callback.isDebug()) {
            callback.printDebug("getChanges", siteUrl, requestUrl.toString(), requestBody);
        }
        ServiceResponseUtil.ChangesHandler handler = new ServiceResponseUtil.ChangesHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, handler);
        return handler.getChanges();
    }

    /**
     * Gets the changes.
     *
     * @param query  the query
     * @param listId the list id
     * @return the changes
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<Change> getChanges(String siteUrl, ChangeQuery query, String listId) throws ServiceException {
        return getChanges(siteUrl, query, listId, null);
    }

    /**
     * Gets the changes.
     *
     * @param query        the query
     * @param listId       the list id
     * @param queryOptions the query options
     * @return the changes
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<Change> getChanges(String siteUrl, ChangeQuery query, String listId, List<IQueryOption> queryOptions) throws ServiceException {
        if (query == null) {
            throw new IllegalArgumentException("query");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }

        /*
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/GetChanges" + queryOptionsString;
        */
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/GetChanges");
        requestUrl.append(sbQuery);
        String requestBody = query.toString();
        if (callback.isDebug()) {
            callback.printDebug("getChanges", siteUrl, requestUrl.toString(), requestBody);
        }
        ServiceResponseUtil.ChangesHandler handler = new ServiceResponseUtil.ChangesHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, handler);
        return handler.getChanges();
    }

    /**
     * Gets the list item changes.
     *
     * @param listId the list id
     * @param query  the query
     * @return the list item changes
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<Change> getListItemChanges(String siteUrl, String listId, ChangeLogItemQuery query) throws ServiceException {
        return getListItemChanges(siteUrl, listId, query, null);
    }

    /**
     * Gets the list item changes.
     *
     * @param listId       the list id
     * @param query        the query
     * @param queryOptions the query options
     * @return the list item changes
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<Change> getListItemChanges(String siteUrl, String listId, ChangeLogItemQuery query, List<IQueryOption> queryOptions) throws ServiceException {
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (query == null) {
            throw new IllegalArgumentException("query");
        }

        /*
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/getlistitemchangessincetoken" + queryOptionsString;
        */
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/getlistitemchangessincetoken");
        requestUrl.append(sbQuery);
        String requestBody = query.toString();
        if (callback.isDebug()) {
            callback.printDebug("getListItemChanges", siteUrl, requestUrl.toString(), requestBody);
        }
        ServiceResponseUtil.ChangesHandler handler = new ServiceResponseUtil.ChangesHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, handler);
        return handler.getChanges();
    }

    /**
     * Render list data.
     *
     * @param listId  the list id
     * @param viewXml the view xml
     * @return the string
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public String renderListData(String siteUrl, String listId, String viewXml) throws ServiceException {
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (viewXml == null) {
            throw new IllegalArgumentException("viewXml");
        }

        // String requestUrl = "_api/web/lists('" + listId + "')/RenderListData(@viewXml)?@viewXml='" + Util.encodeUrl(viewXml) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/RenderListData(@viewXml)?@viewXml='" + Util.encodeUrl(viewXml) + "'");
        if (callback.isDebug()) {
            callback.printDebug("renderListData", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.RenderListDataHandler handler = new ServiceResponseUtil.RenderListDataHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.getRenderListData();
    }

    /**
     * Render list form data.
     *
     * @param listId the list id
     * @param itemId the item id
     * @param formId the form id
     * @param mode   the mode
     * @return the string
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public String renderListFormData(String siteUrl, String listId, int itemId, String formId, ControlMode mode) throws ServiceException {
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (itemId < 0) {
            throw new IllegalArgumentException("The itemId must be non-negative.");
        }

        if (formId == null) {
            throw new IllegalArgumentException("formId");
        }

        // String requestUrl = "_api/web/lists('" + listId + "')/renderlistformdata(itemid=" + itemId + ", formid='" + formId + "', mode=" + mode.getValue() + ")");
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/renderlistformdata(itemid=" + itemId + ", formid='" + formId + "', mode=" + mode.getValue() + ")");
        if (callback.isDebug()) {
            callback.printDebug("renderListFormData", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.RenderListDataHandler handler = new ServiceResponseUtil.RenderListDataHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.getRenderListData();
    }

    /**
     * Reserve list item id.
     *
     * @param listId the list id
     * @return the int
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public int reserveListItemId(String siteUrl, String listId) throws ServiceException {
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/ReserveListItemId");
        if (callback.isDebug()) {
            callback.printDebug("reserveListItemId", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ReserveListItemIdHandler handler = new ServiceResponseUtil.ReserveListItemIdHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.getReserveListItemId();
    }

    /**
     * Reset role inheritance.
     *
     * @param listId the list id
     * @param itemId the item id
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean resetRoleInheritance(String siteUrl, String listId, int itemId) throws ServiceException {
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (itemId < 0) {
            throw new IllegalArgumentException("The itemId must be non-negative.");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/items(" + itemId + ")/ResetRoleInheritance");
        if (callback.isDebug()) {
            callback.printDebug("resetRoleInheritance", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("ResetRoleInheritance");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.isSuccess();
    }

    /**
     * Gets the related fields.
     *
     * @param listId the list id
     * @return the related fields
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<Field> getRelatedFields(String siteUrl, String listId) throws ServiceException {
        return getRelatedFields(siteUrl, listId, null);
    }

    /**
     * Gets the related fields.
     *
     * @param listId       the list id
     * @param queryOptions the query options
     * @return the related fields
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<Field> getRelatedFields(String siteUrl, String listId, List<IQueryOption> queryOptions) throws ServiceException {
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        /*
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/GetRelatedFields" + queryOptionsString;
        */
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/GetRelatedFields");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getRelatedFields", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.FieldsHandler handler = new ServiceResponseUtil.FieldsHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.getFields();
    }

    /**
     * Gets the site templates.
     *
     * @param locale the locale
     * @return the site templates
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    private List<SiteTemplate> getSiteTemplates(String siteUrl, Locale locale) throws ServiceException {
        return getSiteTemplates(siteUrl, locale, false);
    }

    /**
     * Gets the site templates.
     *
     * @param locale       the locale
     * @param queryOptions the query options
     * @return the site templates
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<SiteTemplate> getSiteTemplates(String siteUrl, Locale locale, List<IQueryOption> queryOptions) throws ServiceException {
        return getSiteTemplates(siteUrl, locale, false, queryOptions);
    }

    /**
     * Gets the site templates.
     *
     * @param locale               the locale
     * @param includeCrossLanguage the include cross language
     * @return the site templates
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<SiteTemplate> getSiteTemplates(String siteUrl, Locale locale, boolean includeCrossLanguage) throws ServiceException {
        return getSiteTemplates(siteUrl, locale, includeCrossLanguage, null);
    }

    /**
     * Gets the site templates.
     *
     * @param locale               the locale
     * @param includeCrossLanguage the include cross language
     * @param queryOptions         the query options
     * @return the site templates
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<SiteTemplate> getSiteTemplates(String siteUrl, Locale locale, boolean includeCrossLanguage, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        String includeCrossLanguageParameter = includeCrossLanguage ? ", doincludecrosslanguage=true" : "";
        /*
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/GetAvailableWebTemplates(lcid=" + locale.getValue() + includeCrossLanguageParameter + ")" + queryOptionsString;
        */
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/GetAvailableWebTemplates(lcid=" + locale.getValue() + includeCrossLanguageParameter + ")");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getSiteTemplates", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.SiteTemplatesHandler handler = new ServiceResponseUtil.SiteTemplatesHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.getSiteTemplates();
    }

    /**
     * Gets the effective base permissions.
     *
     * @return the effective base permissions
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<BasePermission.Defines> getEffectiveBasePermissions(String siteUrl) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/EffectiveBasePermissions");
        if (callback.isDebug()) {
            callback.printDebug("getEffectiveBasePermissions", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BasePermissionsHandler handler = new ServiceResponseUtil.BasePermissionsHandler("EffectiveBasePermissions");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.getBasePermissions();
    }

    /**
     * Creates the site.
     *
     * @param siteInfo the site info
     * @return the site
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public Site createSite(String siteUrl, SiteCreationInfo siteInfo) throws ServiceException {
        if (siteInfo == null) {
            throw new IllegalArgumentException("siteInfo");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/webs/Add");
        String requestBody = siteInfo.toString();
        if (callback.isDebug()) {
            callback.printDebug("createSite", siteUrl, requestUrl.toString(), requestBody);
        }
        ServiceResponseUtil.SiteHandler handler = new ServiceResponseUtil.SiteHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, handler);
        return handler.getSite();
    }

    // [Start] 22014: Method to create team site
    public void createTeamSite(ServiceAdmin serviceAdmin,
                               String sUrl,
                               String sTitle,
                               String sDescription,
                               String sLcid,
                               String sHubSiteId,
                               String sClassification,
                               Constant.SiteType type,
                               String[] aOwner)
            throws ServiceException {
        if (serviceAdmin == null) {
            throw new IllegalArgumentException("seviceAdmin");
        }
        if (sTitle == null || sTitle.isEmpty()) {
            sTitle = Util.Url.getFileName(sUrl);
        }
        StringBuilder requestUrl = new StringBuilder("_api/GroupSiteManager/CreateGroupEx");
        SerializableJSONObject jo = new SerializableJSONObject();
        jo.put("alias", Util.Url.getFileName(sUrl));
        jo.put("displayName", sTitle);
        boolean isPublic = false;
        if (Constant.SiteType.PUBLIC.equals(type)) {
            isPublic = true;
        }
        jo.put("isPublic", isPublic);
        SerializableJSONArray jaOwners = new SerializableJSONArray();
        for (String owner : aOwner) {
            // skip because creator will be the owner automatically
            if (serviceAdmin.getUsername().equalsIgnoreCase(owner)) {
                continue;
            }
            jaOwners.put(Util.encodeJson(owner));
        }
        SerializableJSONObject joOptionalParams = new SerializableJSONObject();
        SerializableJSONObject joCreationOptions = new SerializableJSONObject();
        SerializableJSONObject joCreationOptionResult = new SerializableJSONObject();
        SerializableJSONArray jaCreationOptionResult = new SerializableJSONArray();
        jaCreationOptionResult.put("HubSiteId:" + sHubSiteId);
        // 24468: Change to "Lcid", because using "SPSiteLanguage" will create an inaccessible site
        // jaCreationOptionResult.put("SPSiteLanguage:" + sLcid);
        jaCreationOptionResult.put("Lcid:" + sLcid);
        joCreationOptionResult.put("results", jaCreationOptionResult.getObject());
        if (sHubSiteId != null && !"".equals(sHubSiteId)) {
            joCreationOptions.put("HubSiteId", sHubSiteId);
        }
        joOptionalParams.put("Description", sDescription);
        joOptionalParams.put("Classification", sClassification);
        if (jaOwners.length() > 0) {
            SerializableJSONObject joResult = new SerializableJSONObject();
            joResult.put("results", jaOwners.getObject());
            joOptionalParams.put("Owners", joResult.getObject());
        }
        joOptionalParams.put("CreationOptions", joCreationOptionResult.getObject());
        jo.put("optionalParams", joOptionalParams.getObject());
        String requestBody = jo.toString();
        if (callback.isDebug()) {
            callback.printDebug("createSite", serviceAdmin.getSiteUrl(), requestUrl.toString(), requestBody);
        }
        serviceAdmin.doSendRequest(serviceAdmin.getSiteUrl(), "POST", requestUrl.toString(), requestBody, null);
    }
    // [End] 22014

    /**
     * Update site.
     *
     * @param site the site
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public void updateSite(String siteUrl, Site site) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (site == null) {
            throw new IllegalArgumentException("site");
        }

        updateSite(siteUrl, site.toUpdateJSon());
    }

    public void updateSite(String siteUrl, String requestBody) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web");
        if (callback.isDebug()) {
            callback.printDebug("updateSite", siteUrl, requestUrl.toString(), requestBody);
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, "MERGE", null);
    }

    /**
     * Delete site.
     *
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public void deleteSite(String siteUrl) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web");
        if (callback.isDebug()) {
            callback.printDebug("deleteSite", siteUrl, requestUrl.toString());
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), null, "DELETE", null);
    }

    /**
     * Gets the users.
     *
     * @param groupId the group id
     * @return the users
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<User> getUsers(String siteUrl, int groupId) throws ServiceException {
        return getUsers(siteUrl, groupId, null);
    }

    /**
     * Gets the users.
     *
     * @param groupId      the group id
     * @param queryOptions the query options
     * @return the users
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<User> getUsers(String siteUrl, int groupId, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (groupId <= 0) {
            throw new IllegalArgumentException("groupId");
        }

        /*
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/SiteGroups(" + groupId + ")/Users" + queryOptionsString;
        */
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/SiteGroups(" + groupId + ")/Users");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getUsers", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.UsersHandler handler = new ServiceResponseUtil.UsersHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getUsers();
    }

    // [Start] 21984: methods to get and set sharingInfo
    public ShareInfo getListItemItemSharingInformation(String sSiteUrl, String listId, int itemId, List<IQueryOption> queryOptions)
            throws ServiceException {
        /*
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        String sRequestUrl = new StringBuilder("_api/web/Lists('" + listId + "')/GetItemById('" + itemId + "')/GetSharingInformation" + queryOptionsString;
        */
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/Lists('" + listId + "')/GetItemById('" + itemId + "')/GetSharingInformation");
        requestUrl.append(sbQuery);
        ServiceResponseUtil.ListItemShareInfoHandler handler = new ServiceResponseUtil.ListItemShareInfoHandler();
        doSendRequest(sSiteUrl, "POST", requestUrl.toString(), null, handler);
        return handler.getShareInfo();
    }

    public void createListItemShareLink(String sSiteUrl, String listId, int listItemId, String sRequestBody)
            throws ServiceException {
        // String sRequestUrl = new StringBuilder("_api/web/Lists('" + listId + "')/GetItemById('" + listItemId + "')/ShareLink");
        StringBuilder requestUrl = new StringBuilder("_api/web/Lists('" + listId + "')/GetItemById('" + listItemId + "')/ShareLink");
        doSendRequest(sSiteUrl, "POST", requestUrl.toString(), sRequestBody, null);
    }
    // [End] 21984

    /**
     * Gets the user.
     *
     * @param userId  the user id
     * @param groupId the group id
     * @return the user
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public User getUser(String siteUrl, int userId, int groupId) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (userId <= 0) {
            throw new IllegalArgumentException("userId");
        }
        if (groupId <= 0) {
            throw new IllegalArgumentException("groupId");
        }

        // StringBuilder requestUrl = new StringBuilder("_api/web/SiteGroups(" + groupId + ")/Users/GetById(" + userId + ")");
        StringBuilder requestUrl = new StringBuilder("_api/web/SiteGroups(" + groupId + ")/Users/GetById(" + userId + ")");
        if (callback.isDebug()) {
            callback.printDebug("getUser", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.UserHandler handler = new ServiceResponseUtil.UserHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getUser();
    }

    /**
     * Gets the users.
     *
     * @return the users
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<User> getUsers(String siteUrl) throws ServiceException {
        return getUsers(siteUrl, null);
    }

    /**
     * Gets the users.
     *
     * @param queryOptions the query options
     * @return the users
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<User> getUsers(String siteUrl, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        /*
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/SiteUsers" + queryOptionsString;
        */
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/SiteUsers");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getUsers", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.UsersHandler handler = new ServiceResponseUtil.UsersHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getUsers();
    }

    /**
     * Gets the user.
     *
     * @param userId the user id
     * @return the user
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public User getUser(String siteUrl, int userId) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        // StringBuilder requestUrl = new StringBuilder("_api/web/GetUserById(" + userId + ")");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetUserById(" + userId + ")");
        if (callback.isDebug()) {
            callback.printDebug("getUser", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.UserHandler handler = new ServiceResponseUtil.UserHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getUser();
    }

    /**
     * Gets the user groups.
     *
     * @param userId the user id
     * @return the user groups
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<Group> getUserGroups(String siteUrl, int userId) throws ServiceException {
        return getUserGroups(siteUrl, userId, null);
    }

    /**
     * Gets the user groups.
     *
     * @param userId       the user id
     * @param queryOptions the query options
     * @return the user groups
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<Group> getUserGroups(String siteUrl, int userId, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        /*
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/GetUserById(" + userId + ")/Groups" + queryOptionsString;
        */
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/GetUserById(" + userId + ")/Groups");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getUserGroups", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.GroupsHandler handler = new ServiceResponseUtil.GroupsHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getGroups();
    }

    /**
     * Gets the user.
     *
     * @param loginName the login name
     * @param groupId   the group id
     * @return the user
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public User getUser(String siteUrl, String loginName, int groupId) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (loginName == null) {
            throw new IllegalArgumentException("loginName");
        }
        if (groupId <= 0) {
            throw new IllegalArgumentException("groupId");
        }

        // StringBuilder requestUrl = new StringBuilder("_api/web/SiteGroups(" + groupId + ")/Users/GetByLoginName(@v)?@v='" + Util.encodeUrl(loginName) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/SiteGroups(" + groupId + ")/Users/GetByLoginName(@v)?@v='" + Util.encodeUrl(loginName) + "'");
        if (callback.isDebug()) {
            callback.printDebug("getUser", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.UserHandler handler = new ServiceResponseUtil.UserHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getUser();
    }

    /**
     * Gets the user by email.
     *
     * @param email   the email
     * @param groupId the group id
     * @return the user by email
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public User getUserByEmail(String siteUrl, String email, int groupId) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (email == null) {
            throw new IllegalArgumentException("email");
        }
        if (groupId <= 0) {
            throw new IllegalArgumentException("groupId");
        }

        // StringBuilder requestUrl = new StringBuilder("_api/web/SiteGroups(" + groupId + ")/Users/GetByEmail('" + Util.encodeUrl(email) + "')");
        StringBuilder requestUrl = new StringBuilder("_api/web/SiteGroups(" + groupId + ")/Users/GetByEmail('" + Util.encodeUrl(email) + "')");
        if (callback.isDebug()) {
            callback.printDebug("getUserByEmail", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.UserHandler handler = new ServiceResponseUtil.UserHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getUser();
    }

    /**
     * Delete user.
     *
     * @param userId  the user id
     * @param groupId the group id
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean deleteUser(String siteUrl, int userId, int groupId) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (userId <= 0) {
            throw new IllegalArgumentException("userId");
        }
        if (groupId <= 0) {
            throw new IllegalArgumentException("groupId");
        }

        // boolean success = false;
        // InputStream inputStream = null;

        // StringBuilder requestUrl = new StringBuilder("_api/web/SiteGroups(" + groupId + ")/Users/RemoveById(" + userId + ")");
        StringBuilder requestUrl = new StringBuilder("_api/web/SiteGroups(" + groupId + ")/Users/RemoveById(" + userId + ")");
        if (callback.isDebug()) {
            callback.printDebug("deleteUser", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("RemoveById");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.isSuccess();
    }

    /**
     * Delete user.
     *
     * @param loginName the login name
     * @param groupId   the group id
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean deleteUser(String siteUrl, String loginName, int groupId) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (loginName == null) {
            throw new IllegalArgumentException("loginName");
        }
        if (groupId <= 0) {
            throw new IllegalArgumentException("groupId");
        }

        // StringBuilder requestUrl = new StringBuilder("_api/web/SiteGroups(" + groupId + ")/Users/RemoveByLoginName(@v)?@v='" + Util.encodeUrl(loginName) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/SiteGroups(" + groupId + ")/Users/RemoveByLoginName(@v)?@v='" + Util.encodeUrl(loginName) + "'");
        if (callback.isDebug()) {
            callback.printDebug("deleteUser", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("RemoveByLoginName");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.isSuccess();
    }

    /**
     * Gets the groups.
     *
     * @return the groups
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    private List<Group> getGroups(String siteUrl) throws ServiceException {
        return getGroups(siteUrl, null);
    }

    /**
     * Gets the groups.
     *
     * @param queryOptions the query options
     * @return the groups
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<Group> getGroups(String siteUrl, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        /*
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/SiteGroups" + queryOptionsString;
        */
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/SiteGroups");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getGroups", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.GroupsHandler handler = new ServiceResponseUtil.GroupsHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getGroups();
    }

    /**
     * Creates the user.
     *
     * @param user    the user
     * @param groupId the group id
     * @return the user
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public User createUser(String siteUrl, User user, int groupId) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (user == null) {
            throw new IllegalArgumentException("user");
        }
        if (groupId <= 0) {
            throw new IllegalArgumentException("groupId");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/SiteGroups(" + groupId + ")/Users");
        String requestBody = user.toCreateJSon();
        if (callback.isDebug()) {
            callback.printDebug("createUser", siteUrl, requestUrl.toString(), requestBody);
        }
        ServiceResponseUtil.UserHandler handler = new ServiceResponseUtil.UserHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, handler);
        return handler.getUser();
    }

    /**
     * Update user.
     *
     * @param user    the user
     * @param groupId the group id
     * @return the user
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public User updateUser(String siteUrl, User user, int groupId) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (user == null) {
            throw new IllegalArgumentException("user");
        }
        if (user.getLoginName() == null) {
            throw new IllegalArgumentException("LoginName");
        }
        if (groupId <= 0) {
            throw new IllegalArgumentException("groupId");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/SiteGroups(" + groupId + ")/Users(@v)?@v='" + user.getLoginName() + "'");
        String requestBody = user.toUpdateJSon();
        if (callback.isDebug()) {
            callback.printDebug("updateUser", siteUrl, requestUrl.toString(), requestBody);
        }
        ServiceResponseUtil.UserHandler handler = new ServiceResponseUtil.UserHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, "MERGE", handler);
        return handler.getUser();
    }

    /**
     * Creates the role.
     *
     * @param role the role
     * @return the role
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public Role createRole(String siteUrl, Role role) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (role == null) {
            throw new IllegalArgumentException("role");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/RoleDefinitions");
        String requestBody = role.toCreateJSon();
        if (callback.isDebug()) {
            callback.printDebug("createRole", siteUrl, requestUrl.toString(), requestBody);
        }
        ServiceResponseUtil.RoleHandler handler = new ServiceResponseUtil.RoleHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, handler);
        return handler.getRole();
    }

    /**
     * Delete role.
     *
     * @param roleId the role id
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public void deleteRole(String siteUrl, int roleId) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (roleId <= 0) {
            throw new IllegalArgumentException("roleId");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/RoleDefinitions(" + roleId + ")");
        if (callback.isDebug()) {
            callback.printDebug("deleteRole", siteUrl, requestUrl.toString());
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), null, "DELETE", null);
    }

    /**
     * Update role.
     *
     * @param role the role
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public void updateRole(String siteUrl, Role role) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (role == null) {
            throw new IllegalArgumentException("role");
        }
        if (role.getId() == null) {
            throw new IllegalArgumentException("Invalid Role Id");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/RoleDefinitions(" + role.getId() + ")");
        String requestBody = role.toUpdateJSon();
        if (callback.isDebug()) {
            callback.printDebug("updateRole", siteUrl, requestUrl.toString(), requestBody);
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, "MERGE", null);
    }

    /**
     * Update role.
     *
     * @param RoleId      the destination role Id
     * @param requestBody the update request json
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public void updateRole(String siteUrl, String RoleId, String requestBody) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (RoleId == null) {
            throw new IllegalArgumentException("role");
        }
        if (requestBody == null) {
            throw new IllegalArgumentException("newName");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/RoleDefinitions(" + RoleId + ")");
        if (callback.isDebug()) {
            callback.printDebug("updateRole", siteUrl, requestUrl.toString(), requestBody);
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, "MERGE", null);
    }

    /**
     * Creates the group.
     *
     * @param group the group
     * @return the group
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public Group createGroup(String siteUrl, Group group) throws ServiceException {
        if (group == null) {
            throw new IllegalArgumentException("group");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/SiteGroups");
        String requestBody = group.toCreateJSon();
        if (callback.isDebug()) {
            callback.printDebug("createGroup", siteUrl, requestUrl.toString(), requestBody);
        }
        ServiceResponseUtil.GroupHandler handler = new ServiceResponseUtil.GroupHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, handler);
        return handler.getGroup();
    }

    /**
     * Update group.
     *
     * @param group the group
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public void updateGroup(String siteUrl, Group group) throws ServiceException {
        if (group == null) {
            throw new IllegalArgumentException("group");
        }
        if (group.getId() == 0) {
            throw new IllegalArgumentException("Id");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/SiteGroups(" + group.getId() + ")");
        String requestBody = group.toUpdateJSon();
        if (callback.isDebug()) {
            callback.printDebug("updateGroup", siteUrl, requestUrl.toString(), requestBody);
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, "MERGE", null);
    }

    public void updateGroup(String siteUrl, int groupId, String requestBody) throws ServiceException {
        if (groupId <= 0) {
            throw new IllegalArgumentException("groupId");
        }
        if (requestBody == null) {
            throw new IllegalArgumentException("requestBody");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/SiteGroups(" + groupId + ")");
        if (callback.isDebug()) {
            callback.printDebug("updateGroup", siteUrl, requestUrl.toString(), requestBody);
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, "MERGE", null);
    }

    /**
     * Delete group.
     *
     * @param groupId the group id
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean deleteGroup(String siteUrl, int groupId) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/SiteGroups/RemoveById(" + groupId + ")");
        if (callback.isDebug()) {
            callback.printDebug("deleteGroup", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("RemoveById");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.isSuccess();
    }

    /**
     * Delete group.
     *
     * @param loginName the login name
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean deleteGroup(String siteUrl, String loginName) throws ServiceException {
        if (loginName == null) {
            throw new IllegalArgumentException("loginName");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/SiteGroups/RemoveByLoginName('" + Util.encodeUrl(loginName) + "')");
        if (callback.isDebug()) {
            callback.printDebug("deleteGroup", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("RemoveByLoginName");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.isSuccess();
    }

    /**
     * Gets the group.
     *
     * @param groupId the group id
     * @return the group
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public Group getGroup(String siteUrl, int groupId, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        /*
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/SiteGroups(" + groupId + ")" + queryOptionsString;
        */
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/SiteGroups(" + groupId + ")");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getGroup", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.GroupHandler handler = new ServiceResponseUtil.GroupHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getGroup();
    }

    /**
     * Gets the group owner.
     *
     * @param groupId the group id
     * @return the group owner
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public User getGroupOwner(String siteUrl, int groupId) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/SiteGroups(" + groupId + ")/Owner");
        if (callback.isDebug()) {
            callback.printDebug("getGroupOwner", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.UserHandler handler = new ServiceResponseUtil.UserHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getUser();
    }

    /**
     * Gets the group users.
     *
     * @param groupId the group id
     * @return the group users
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<User> getGroupUsers(String siteUrl, int groupId) throws ServiceException {
        return getGroupUsers(siteUrl, groupId, null);
    }

    /**
     * Gets the group users.
     *
     * @param groupId      the group id
     * @param queryOptions the query options
     * @return the group users
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<User> getGroupUsers(String siteUrl, int groupId, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        /*
        InputStream inputStream = null;
        List<User> users = new ArrayList<User>();
        
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/SiteGroups(" + groupId + ")/Users" + queryOptionsString;
        */
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/SiteGroups(" + groupId + ")/Users");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getGroupUsers", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.UsersHandler handler = new ServiceResponseUtil.UsersHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getUsers();
    }

    /**
     * Gets the group.
     *
     * @param loginName the login name
     * @return the group
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public Group getGroup(String siteUrl, String loginName, List<IQueryOption> queryOptions) throws ServiceException {
        if (loginName == null) {
            throw new IllegalArgumentException("loginName");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        /*
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/SiteGroups/GetByName('" + Util.encodeUrl(loginName) + "')" + queryOptionsString;
        */
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/SiteGroups/GetByName('" + Util.encodeUrl(loginName) + "')");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getGroup", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.GroupHandler handler = new ServiceResponseUtil.GroupHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getGroup();
    }

    /**
     * Gets the group owner.
     *
     * @param loginName the login name
     * @return the group owner
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public User getGroupOwner(String siteUrl, String loginName) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (loginName == null) {
            throw new IllegalArgumentException("loginName");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/SiteGroups/GetByName('" + Util.encodeUrl(loginName) + "')/Owner");
        if (callback.isDebug()) {
            callback.printDebug("getGroupOwner", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.UserHandler handler = new ServiceResponseUtil.UserHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getUser();
    }

    /**
     * Gets the group users.
     *
     * @param loginName the login name
     * @return the group users
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<User> getGroupUsers(String siteUrl, String loginName) throws ServiceException {
        return getGroupUsers(siteUrl, loginName, null);
    }

    /**
     * Gets the group users.
     *
     * @param loginName    the login name
     * @param queryOptions the query options
     * @return the group users
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<User> getGroupUsers(String siteUrl, String loginName, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (loginName == null) {
            throw new IllegalArgumentException("loginName");
        }
        /*
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/SiteGroups/GetByName('" + Util.encodeUrl(loginName) + "')/Users" + queryOptionsString;
        */
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/SiteGroups/GetByName('" + Util.encodeUrl(loginName) + "')/Users");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getGroupUsers", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.UsersHandler handler = new ServiceResponseUtil.UsersHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getUsers();
    }

    /**
     * Gets the role assignments.
     *
     * @return the role assignments
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    /*
    public List<Integer> getRoleAssignments(String siteUrl) throws ServiceException {
        return getRoleAssignments(siteUrl, null);
    }
    */

    /**
     * Gets the role assignments.
     *
     * @param queryOptions the query options
     * @return the role assignments
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    /*
    public List<Integer> getRoleAssignments(String siteUrl, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        InputStream inputStream = null;
        List<Integer> list = new ArrayList<Integer>();

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/roleassignments" + queryOptionsString;
        if (callback.isDebug()) {
            callback.printDebug("getRoleAssignments", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.RoleAssignmentsHandler handler = new ServiceResponseUtil.RoleAssignmentsHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getRoleAssignments();
    }
    */

    /**
     * Gets the role assignments.
     *
     * @return the role assignments
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<RoleAssignment> getRoleAssignments(String siteUrl, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/roleassignments");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getRoleAssignments", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.RoleAssignmentsHandler handler = new ServiceResponseUtil.RoleAssignmentsHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getRoleAssignments();
    }

    // [Start] 24258: Support to get single role assignment by id
    public RoleAssignment getRoleAssignment(String siteUrl, int iPrincipalId, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        // StringBuilder requestUrl = new StringBuilder("_api/web/roleassignments/GetByPrincipalId('" + iPrincipalId + "')" + Util.queryOptionsToString(queryOptions);
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/roleassignments/GetByPrincipalId('" + iPrincipalId + "')");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getRoleAssignment", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.RoleAssignmentHandler handler = new ServiceResponseUtil.RoleAssignmentHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getRoleAssignment();
    }
    // [End] 24258

    /**
     * Adds the role assignment.
     *
     * @param principalId the principal id
     * @param roleId      the role id
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean addRoleAssignment(String siteUrl, int principalId, int roleId) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/roleassignments/addroleassignment(principalid=" + principalId + ",roledefid=" + roleId + ")");
        if (callback.isDebug()) {
            callback.printDebug("addRoleAssignment", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("AddRoleAssignment");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.isSuccess();
    }

    /**
     * Removes the role assignment.
     *
     * @param principalId the principal id
     * @param roleId      the role id
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean removeRoleAssignment(String siteUrl, int principalId, int roleId) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/roleassignments/removeroleassignment(principalid=" + principalId + ",roledefid=" + roleId + ")");
        if (callback.isDebug()) {
            callback.printDebug("removeRoleAssignment", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("RemoveRoleAssignment");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.isSuccess();
    }

    public boolean removeListRoleAssignment(String siteUrl, String listId, int principalId, int roleId) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/roleassignments/removeroleassignment(principalid=" + principalId + ",roledefid=" + roleId + ")");
        if (callback.isDebug()) {
            callback.printDebug("removeRoleAssignment", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("RemoveRoleAssignment");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.isSuccess();
    }

    /**
     * Gets the regional settings.
     *
     * @return the regional settings
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public RegionalSettings getRegionalSettings(String siteUrl) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/RegionalSettings");
        if (callback.isDebug()) {
            callback.printDebug("getRegionalSettings", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.RegionalSettingsHandler handler = new ServiceResponseUtil.RegionalSettingsHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getRegionalSettings();
    }

    /**
     * Gets the time zones.
     *
     * @return the time zones
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<TimeZone> getTimeZones(String siteUrl) throws ServiceException {
        return getTimeZones(siteUrl, null);
    }

    /**
     * Gets the time zones.
     *
     * @param queryOptions the query options
     * @return the time zones
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<TimeZone> getTimeZones(String siteUrl, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/RegionalSettings/TimeZones");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getTimeZones", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.TimeZonesHandler handler = new ServiceResponseUtil.TimeZonesHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getTimeZones();
    }

    /**
     * Gets the time zone.
     *
     * @return the time zone
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public TimeZone getTimeZone(String siteUrl) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/RegionalSettings/TimeZone");
        if (callback.isDebug()) {
            callback.printDebug("getTimeZone", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.TimeZoneHandler handler = new ServiceResponseUtil.TimeZoneHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getTimeZone();
    }

    /**
     * Gets the theme info.
     *
     * @return the theme info
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public ThemeInfo getThemeInfo(String siteUrl) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/ThemeInfo");
        if (callback.isDebug()) {
            callback.printDebug("getThemeInfo", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ThemeInfoHandler handler = new ServiceResponseUtil.ThemeInfoHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getThemeInfo();
    }

    /**
     * Gets the workflow templates.
     *
     * @return the workflow templates
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<WorkflowTemplate> getWorkflowTemplates(String siteUrl) throws ServiceException {
        return getWorkflowTemplates(siteUrl, null);
    }

    /**
     * Gets the workflow templates.
     *
     * @param queryOptions the query options
     * @return the workflow templates
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<WorkflowTemplate> getWorkflowTemplates(String siteUrl, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/WorkflowTemplates");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getWorkflowTemplates", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.WorkflowTemplatesHandler handler = new ServiceResponseUtil.WorkflowTemplatesHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getWorkflowTemplates();
    }

    // [Start] 23322: Added options to get site user info list 

    /**
     * Gets the user info list.
     *
     * @return the user info list
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public com.independentsoft.share.List getUserInfoList(String siteUrl)
            throws ServiceException {
        /*
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/SiteUserInfoList");
        if (callback.isDebug()) {
            callback.printDebug("getUserInfoList", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ListHandler handler = new ServiceResponseUtil.ListHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getList();
        */
        return getUserInfoList(siteUrl, new ArrayList<IQueryOption>());
    }

    public com.independentsoft.share.List getUserInfoList(String siteUrl, List<IQueryOption> queryOptions)
            throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/SiteUserInfoList");
        requestUrl.append(sbQuery);

        if (callback.isDebug()) {
            callback.printDebug("getUserInfoList", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ListHandler handler = new ServiceResponseUtil.ListHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getList();
    }
    // [End] 23322

    /**
     * Gets the root folder.
     *
     * @return the root folder
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public Folder getRootFolder(String siteUrl) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/RootFolder");
        if (callback.isDebug()) {
            callback.printDebug("getRootFolder", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.FolderHandler handler = new ServiceResponseUtil.FolderHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getFolder();
    }

    /**
     * Gets the top navigation bar.
     *
     * @return the top navigation bar
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    private List<NavigationNode> getTopNavigationBar(String siteUrl) throws ServiceException {
        return getTopNavigationBar(siteUrl, null);
    }

    /**
     * Gets the top navigation bar.
     *
     * @param queryOptions the query options
     * @return the top navigation bar
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<NavigationNode> getTopNavigationBar(String siteUrl, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/Navigation/TopNavigationBar");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getTopNavigationBar", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.NavigationNodesHandler handler = new ServiceResponseUtil.NavigationNodesHandler(siteUrl, -1);
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getNavigationNodes();
    }

    /**
     * Gets the quick launch.
     *
     * @return the quick launch
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    private List<NavigationNode> getQuickLaunch(String siteUrl) throws ServiceException {
        return getQuickLaunch(siteUrl, null);
    }

    /**
     * Gets the quick launch.
     *
     * @param queryOptions the query options
     * @return the quick launch
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<NavigationNode> getQuickLaunch(String siteUrl, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/Navigation/QuickLaunch");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getQuickLaunch", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.NavigationNodesHandler handler = new ServiceResponseUtil.NavigationNodesHandler(siteUrl, -1);
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getNavigationNodes();
    }

    /**
     * Updates the navigation node.
     *
     * @param id   the id of node that will be updated
     * @param node the node used to update the node having that id
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public void updateNavigationNode(String siteUrl, int id, NavigationNode node) throws ServiceException {
        if (id <= 0) {
            throw new IllegalArgumentException("The parameter id must be non-negative.");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/Navigation/GetNodeById(" + id + ")");
        String requestBody = node.toCreateJSon();
        if (callback.isDebug()) {
            callback.printDebug("updateNavigationNode", siteUrl, requestUrl.toString());
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, "MERGE", null);
    }

    /**
     * Gets the navigation node.
     *
     * @param id the id
     * @return the navigation node
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public NavigationNode getNavigationNode(String siteUrl, int id) throws ServiceException {
        if (id <= 0) {
            throw new IllegalArgumentException("The parameter id must be non-negative.");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/Navigation/GetNodeById(" + id + ")");
        if (callback.isDebug()) {
            callback.printDebug("getNavigationNode", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.NavigationNodeHandler handler = new ServiceResponseUtil.NavigationNodeHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getNavigationNode();
    }

    /**
     * Gets the navigation node children.
     *
     * @param id the id
     * @return the navigation node children
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    private List<NavigationNode> getNavigationNodeChildren(String siteUrl, int id) throws ServiceException {
        return getNavigationNodeChildren(siteUrl, id, null);
    }

    /**
     * Gets the navigation node children.
     *
     * @param id           the id
     * @param queryOptions the query options
     * @return the navigation node children
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<NavigationNode> getNavigationNodeChildren(String siteUrl, int id, List<IQueryOption> queryOptions) throws ServiceException {
        if (id <= 0) {
            throw new IllegalArgumentException("Invalid node id");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/Navigation/GetNodeById(" + id + ")/Children");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getNavigationNodeChildren", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.NavigationNodesHandler handler = new ServiceResponseUtil.NavigationNodesHandler(siteUrl, id);
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getNavigationNodes();
    }

    /**
     * Checks if is shared navigation.
     *
     * @return true, if is shared navigation
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean isSharedNavigation(String siteUrl) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/Navigation");
        if (callback.isDebug()) {
            callback.printDebug("isSharedNavigation", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("UseShared");
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.isSuccess();
    }

    /**
     * Gets the recycle bin items.
     *
     * @return the recycle bin items
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<RecycleBinItem> getRecycleBinItems(String siteUrl) throws ServiceException {
        return getRecycleBinItems(siteUrl, null);
    }

    /**
     * Gets the recycle bin items.
     *
     * @param queryOptions the query options
     * @return the recycle bin items
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<RecycleBinItem> getRecycleBinItems(String siteUrl, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/RecycleBin");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getRecycleBinItems", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.RecycleBinItemsHandler handler = new ServiceResponseUtil.RecycleBinItemsHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getRecycleBinItems();
    }

    /**
     * Gets the recycle bin item.
     *
     * @param id the id
     * @return the recycle bin item
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public RecycleBinItem getRecycleBinItem(String siteUrl, String id) throws ServiceException {
        if (id == null) {
            throw new IllegalArgumentException("id");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/RecycleBin('" + id + "')");
        if (callback.isDebug()) {
            callback.printDebug("getRecycleBinItem", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.RecycleBinItemHandler handler = new ServiceResponseUtil.RecycleBinItemHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getRecycleBinItem();
    }

    /**
     * Gets the recycle bin item author.
     *
     * @param id the id
     * @return the recycle bin item author
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public User getRecycleBinItemAuthor(String siteUrl, String id) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (id == null) {
            throw new IllegalArgumentException("id");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/RecycleBin('" + id + "')/Author");
        if (callback.isDebug()) {
            callback.printDebug("getRecycleBinItemAuthor", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.UserHandler handler = new ServiceResponseUtil.UserHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getUser();
    }

    /**
     * Gets the recycle bin item deleted by.
     *
     * @param id the id
     * @return the recycle bin item deleted by
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public User getRecycleBinItemDeletedBy(String siteUrl, String id) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (id == null) {
            throw new IllegalArgumentException("id");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/RecycleBin('" + id + "')/DeletedBy");
        if (callback.isDebug()) {
            callback.printDebug("getRecycleBinItemDeletedBy", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.UserHandler handler = new ServiceResponseUtil.UserHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getUser();
    }

    /**
     * Gets the user custom actions.
     *
     * @return the user custom actions
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<Role> getUserCustomActions(String siteUrl) throws ServiceException {
        return getUserCustomActions(siteUrl, null);
    }

    /**
     * Gets the user custom actions.
     *
     * @param queryOptions the query options
     * @return the user custom actions
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<Role> getUserCustomActions(String siteUrl, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/UserCustomActions");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getUserCustomActions", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.RolesHandler handler = new ServiceResponseUtil.RolesHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getRoles();
    }

    /**
     * Gets the roles.
     *
     * @return the roles
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    private List<Role> getRoles(String siteUrl) throws ServiceException {
        return getRoles(siteUrl, null);
    }

    /**
     * Gets the roles.
     *
     * @param queryOptions the query options
     * @return the roles
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<Role> getRoles(String siteUrl, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/RoleDefinitions");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getRoles", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.RolesHandler handler = new ServiceResponseUtil.RolesHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getRoles();
    }

    /**
     * Gets the role.
     *
     * @param roleId the role id
     * @return the role
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public Role getRole(String siteUrl, int roleId) throws ServiceException {
        if (roleId <= 0) {
            throw new IllegalArgumentException("The parameter roleId must be non-negative.");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/RoleDefinitions(" + roleId + ")");
        if (callback.isDebug()) {
            callback.printDebug("getRole", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.RoleHandler handler = new ServiceResponseUtil.RoleHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getRole();
    }

    /**
     * Gets the role.
     *
     * @param name the name
     * @return the role
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public Role getRole(String siteUrl, String name) throws ServiceException {
        if (name == null) {
            throw new IllegalArgumentException("name");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/RoleDefinitions/GetByName('" + Util.encodeUrl(name) + "')");
        if (callback.isDebug()) {
            callback.printDebug("getRole", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.RoleHandler handler = new ServiceResponseUtil.RoleHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getRole();
    }

    /**
     * Gets the role.
     *
     * @param type the type
     * @return the role
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public Role getRole(String siteUrl, RoleType type) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/RoleDefinitions/GetByType('" + type.getValue() + "')");
        if (callback.isDebug()) {
            callback.printDebug("getRole", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.RoleHandler handler = new ServiceResponseUtil.RoleHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getRole();
    }

    /**
     * Gets the field.
     *
     * @param id the id
     * @return the field
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public Field getField(String siteUrl, String id, List<IQueryOption> queryOptions) throws ServiceException {
        if (id == null) {
            throw new IllegalArgumentException("id");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/fields('" + id + "')");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getField", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.FieldHandler handler = new ServiceResponseUtil.FieldHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getField();
    }

    /**
     * Gets the field by title.
     *
     * @param title the title
     * @return the field by title
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public Field getFieldByTitle(String siteUrl, String title) throws ServiceException {
        if (title == null) {
            throw new IllegalArgumentException("title");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/fields/getbytitle('" + title + "')");
        if (callback.isDebug()) {
            callback.printDebug("getFieldByTitle", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.FieldHandler handler = new ServiceResponseUtil.FieldHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getField();
    }

    /**
     * Gets the field by internal name or title.
     *
     * @param name the name
     * @return the field by internal name or title
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public Field getFieldByInternalNameOrTitle(String siteUrl, String name) throws ServiceException {
        if (name == null) {
            throw new IllegalArgumentException("name");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/fields/getbyinternalnameortitle('" + name + "')");
        if (callback.isDebug()) {
            callback.printDebug("getFieldByInternalNameOrTitle", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.FieldHandler handler = new ServiceResponseUtil.FieldHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getField();
    }

    public void deleteField(String siteUrl, String id) throws ServiceException {
        if (id == null) {
            throw new IllegalArgumentException("id");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/fields('" + id + "')");
        if (callback.isDebug()) {
            callback.printDebug("deleteField", siteUrl, requestUrl.toString());
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), null, "DELETE", null);
    }

    /**
     * Sets the show in display form.
     *
     * @param id     the id
     * @param listId the list id
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean setShowInDisplayForm(String siteUrl, String id, String listId) throws ServiceException {
        return setShowInDisplayForm(siteUrl, id, listId, true);
    }

    /**
     * Sets the show in display form.
     *
     * @param id     the id
     * @param listId the list id
     * @param show   the show
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean setShowInDisplayForm(String siteUrl, String id, String listId, boolean show) throws ServiceException {
        if (id == null) {
            throw new IllegalArgumentException("id");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/fields('" + id + "')/setshowindisplayform(" + Boolean.toString(show).toLowerCase() + ")");
        if (callback.isDebug()) {
            callback.printDebug("setShowInDisplayForm", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("SetShowInDisplayForm");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.isSuccess();
    }

    /**
     * Sets the show in edit form.
     *
     * @param id     the id
     * @param listId the list id
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean setShowInEditForm(String siteUrl, String id, String listId) throws ServiceException {
        return setShowInEditForm(siteUrl, id, listId, true);
    }

    /**
     * Sets the show in edit form.
     *
     * @param id     the id
     * @param listId the list id
     * @param show   the show
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean setShowInEditForm(String siteUrl, String id, String listId, boolean show) throws ServiceException {
        if (id == null) {
            throw new IllegalArgumentException("id");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/fields('" + id + "')/setshowineditform(" + Boolean.toString(show).toLowerCase() + ")");
        if (callback.isDebug()) {
            callback.printDebug("setShowInEditForm", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("SetShowInEditForm");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.isSuccess();
    }

    /**
     * Sets the show in new form.
     *
     * @param id     the id
     * @param listId the list id
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean setShowInNewForm(String siteUrl, String id, String listId) throws ServiceException {
        return setShowInNewForm(siteUrl, id, listId, true);
    }

    /**
     * Sets the show in new form.
     *
     * @param id     the id
     * @param listId the list id
     * @param show   the show
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean setShowInNewForm(String siteUrl, String id, String listId, boolean show) throws ServiceException {
        if (id == null) {
            throw new IllegalArgumentException("id");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/fields('" + id + "')/setshowinnewform(" + Boolean.toString(show).toLowerCase() + ")");
        if (callback.isDebug()) {
            callback.printDebug("setShowInNewForm", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("SetShowInNewForm");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.isSuccess();
    }

    /**
     * Gets the field.
     *
     * @param id     the id
     * @param listId the list id
     * @return the field
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public Field getField(String siteUrl, String id, String listId) throws ServiceException {
        if (id == null) {
            throw new IllegalArgumentException("id");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/fields('" + id + "')");
        if (callback.isDebug()) {
            callback.printDebug("getField", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.FieldHandler handler = new ServiceResponseUtil.FieldHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getField();
    }

    /**
     * Gets the field by title.
     *
     * @param title  the title
     * @param listId the list id
     * @return the field by title
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public Field getFieldByTitle(String siteUrl, String title, String listId) throws ServiceException {
        if (title == null) {
            throw new IllegalArgumentException("title");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/fields/getbytitle('" + title + "')");
        if (callback.isDebug()) {
            callback.printDebug("getFieldByTitle", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.FieldHandler handler = new ServiceResponseUtil.FieldHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getField();
    }

    /**
     * Gets the field by internal name or title.
     *
     * @param name   the name
     * @param listId the list id
     * @return the field by internal name or title
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public Field getFieldByInternalNameOrTitle(String siteUrl, String name, String listId) throws ServiceException {
        if (name == null) {
            throw new IllegalArgumentException("name");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/fields/getbyinternalnameortitle('" + name + "')");
        if (callback.isDebug()) {
            callback.printDebug("getFieldByInternalNameOrTitle", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.FieldHandler handler = new ServiceResponseUtil.FieldHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getField();
    }

    /**
     * Creates the field.
     *
     * @param field the field
     * @return the field
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public Field createField(String siteUrl, FieldSchemaXml field) throws ServiceException {
        if (field == null) {
            throw new IllegalArgumentException("field");
        }
        return createField(siteUrl, field.toString());
    }

    public Field createField(String siteUrl, String sSchemaXml) throws ServiceException {
        if (sSchemaXml == null) {
            throw new IllegalArgumentException("sSchemaXml");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/fields/createfieldasxml");
        StringBuilder requestBody = new StringBuilder("{ 'parameters': { '__metadata': { 'type': 'SP.XmlSchemaFieldCreationInformation' }, 'SchemaXml': '");
        requestBody.append(sSchemaXml);
        requestBody.append("' } }");
        if (callback.isDebug()) {
            callback.printDebug("createField", siteUrl, requestUrl.toString(), requestBody.toString());
        }
        ServiceResponseUtil.FieldHandler handler = new ServiceResponseUtil.FieldHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody.toString(), handler);
        return handler.getField();
    }

    /**
     * Adds the dependent lookup field.
     *
     * @param displayName          the display name
     * @param primaryLookupFieldId the primary lookup field id
     * @param showField            the show field
     * @param listId               the list id
     * @return the field
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public Field addDependentLookupField(String siteUrl, String displayName, String primaryLookupFieldId, String showField, String listId) throws ServiceException {
        if (displayName == null) {
            throw new IllegalArgumentException("displayName");
        }
        if (primaryLookupFieldId == null) {
            throw new IllegalArgumentException("primaryLookupFieldId");
        }
        if (showField == null) {
            throw new IllegalArgumentException("showField");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/fields/adddependentlookupfield(displayname='" + Util.encodeUrl(displayName) + "', primarylookupfieldid='" + Util.encodeUrl(primaryLookupFieldId) + "', showfield='" + Util.encodeUrl(showField) + "')");
        if (callback.isDebug()) {
            callback.printDebug("addDependentLookupField", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.FieldHandler handler = new ServiceResponseUtil.FieldHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.getField();
    }

    public void updateField(String siteUrl, String fieldId, String requestBody) throws ServiceException {
        if (fieldId == null) {
            throw new IllegalArgumentException("fieldId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/fields('" + fieldId + "')");
        if (callback.isDebug()) {
            callback.printDebug("updateField", siteUrl, requestUrl.toString(), requestBody);
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, "MERGE", null);
    }

    /**
     * Update field.
     *
     * @param fieldId the field id
     * @param listId  the list id
     * @return the field
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public void updateListField(String siteUrl, String fieldId, String listId, String requestBody) throws ServiceException {
        if (fieldId == null) {
            throw new IllegalArgumentException("fieldId");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/fields('" + fieldId + "')");
        if (callback.isDebug()) {
            callback.printDebug("updateField", siteUrl, requestUrl.toString(), requestBody);
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, "MERGE", null);
    }

    /**
     * Creates the field.
     *
     * @param field the field
     * @return the field
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public Field createField(String siteUrl, Field field) throws ServiceException {
        if (field == null) {
            throw new IllegalArgumentException("field");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/fields");
        String requestBody = field.toCreateJSon();
        if (callback.isDebug()) {
            callback.printDebug("createField", siteUrl, requestUrl.toString(), requestBody);
        }
        ServiceResponseUtil.FieldHandler handler = new ServiceResponseUtil.FieldHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, handler);
        return handler.getField();
    }

    /**
     * Creates the field.
     *
     * @param field  the field
     * @param listId the list id
     * @return the field
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public Field createListField(String siteUrl, FieldSchemaXml field, String listId) throws ServiceException {
        if (field == null) {
            throw new IllegalArgumentException("field");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/fields/createfieldasxml");
        StringBuilder requestBody = new StringBuilder("{ 'parameters': { '__metadata': { 'type': 'SP.XmlSchemaFieldCreationInformation' }, 'SchemaXml': '");
        requestBody.append(field.toString());
        requestBody.append("' } }");
        if (callback.isDebug()) {
            callback.printDebug("createListField", siteUrl, requestUrl.toString(), requestBody.toString());
        }
        ServiceResponseUtil.FieldHandler handler = new ServiceResponseUtil.FieldHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody.toString(), handler);
        return handler.getField();
    }

    /**
     * Creates the field.
     *
     * @param field  the field
     * @param listId the list id
     * @return the field
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public Field createListField(String siteUrl, FieldCreationInfo field, String listId) throws ServiceException {
        if (field == null) {
            throw new IllegalArgumentException("field");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/fields/addfield");
        String requestBody = field.toCreateJSon();
        if (callback.isDebug()) {
            callback.printDebug("createListField", siteUrl, requestUrl.toString(), requestBody);
        }
        ServiceResponseUtil.FieldHandler handler = new ServiceResponseUtil.FieldHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, handler);
        return handler.getField();
    }

    /**
     * Creates the field.
     *
     * @param field  the field
     * @param listId the list id
     * @return the field
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public Field createListField(String siteUrl, Field field, String listId) throws ServiceException {
        if (field == null) {
            throw new IllegalArgumentException("field");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/fields");
        String requestBody = field.toCreateJSon();
        if (callback.isDebug()) {
            callback.printDebug("createListField", siteUrl, requestUrl.toString(), requestBody);
        }
        ServiceResponseUtil.FieldHandler handler = new ServiceResponseUtil.FieldHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, handler);
        return handler.getField();
    }

    /**
     * Delete field.
     *
     * @param fieldId the id
     * @param listId  the list id
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public void deleteListField(String siteUrl, String fieldId, String listId) throws ServiceException {
        if (fieldId == null) {
            throw new IllegalArgumentException("fieldId");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/fields('" + fieldId + "')");
        if (callback.isDebug()) {
            callback.printDebug("deleteListField", siteUrl, requestUrl.toString());
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), null, "DELETE", null);
    }

    /**
     * Gets the features.
     *
     * @return the features
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<Feature> getWebFeatures(String siteUrl) throws ServiceException {
        return getWebFeatures(siteUrl, null);
    }

    /**
     * Gets the features.
     *
     * @param queryOptions the query options
     * @return the features
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<Feature> getWebFeatures(String siteUrl, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/Features");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getWebFeatures", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.FeaturesHandler handler = new ServiceResponseUtil.FeaturesHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getFeatures();
    }

    /**
     * Gets the fields.
     *
     * @return the fields
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    private List<Field> getFields(String siteUrl) throws ServiceException {
        return getFields(siteUrl, null);
    }

    /**
     * Gets the fields.
     *
     * @param queryOptions the query options
     * @return the fields
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<Field> getFields(String siteUrl, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/Fields");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getFields", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.FieldsHandler handler = new ServiceResponseUtil.FieldsHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getFields();
    }

    /**
     * Gets the available fields.
     *
     * @return the available fields
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    private List<Field> getAvailableFields(String siteUrl) throws ServiceException {
        return getAvailableFields(siteUrl, null);
    }

    /**
     * Gets the available fields.
     *
     * @param queryOptions the query options
     * @return the available fields
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<Field> getAvailableFields(String siteUrl, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/AvailableFields");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getAvailableFields", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.FieldsHandler handler = new ServiceResponseUtil.FieldsHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getFields();
    }

    /**
     * Gets the list fields.
     *
     * @param listId the list id
     * @return the list fields
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    private List<Field> getListFields(String siteUrl, String listId) throws ServiceException {
        return getListFields(siteUrl, listId, null);
    }

    /**
     * Gets the list fields.
     *
     * @param listId       the list id
     * @param queryOptions the query options
     * @return the list fields
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<Field> getListFields(String siteUrl, String listId, List<IQueryOption> queryOptions) throws ServiceException {
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/Fields");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getListFields", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.FieldsHandler handler = new ServiceResponseUtil.FieldsHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getFields();
    }

    public Field getListField(String siteUrl, String listId, String fieldId, List<IQueryOption> queryOptions) throws ServiceException {
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/Fields('" + Util.encodeUrl(fieldId) + "')");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getListFields", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.FieldHandler handler = new ServiceResponseUtil.FieldHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getField();
    }

    /**
     * Gets the list forms.
     *
     * @param listId the list id
     * @return the list forms
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<Form> getListForms(String siteUrl, String listId) throws ServiceException {
        return getListForms(siteUrl, listId, null);
    }

    /**
     * Gets the list forms.
     *
     * @param listId       the list id
     * @param queryOptions the query options
     * @return the list forms
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<Form> getListForms(String siteUrl, String listId, List<IQueryOption> queryOptions) throws ServiceException {
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/Forms");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getListForms", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ListFormsHandler handler = new ServiceResponseUtil.ListFormsHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getListForms();
    }

    /**
     * Gets the list form.
     *
     * @param listId the list id
     * @param formId the form id
     * @return the list form
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public Form getListForm(String siteUrl, String listId, String formId) throws ServiceException {
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (formId == null) {
            throw new IllegalArgumentException("formId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/Forms('" + formId + "')");
        if (callback.isDebug()) {
            callback.printDebug("getListForm", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ListFormHandler handler = new ServiceResponseUtil.ListFormHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getListForm();
    }

    private ContentType getListContentType(String siteUrl, String listId, String contentTypeId) throws ServiceException {
        return getListContentType(siteUrl, listId, contentTypeId, null);
    }

    public ContentType getListContentType(String siteUrl, String listId, String contentTypeId, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (contentTypeId == null) {
            throw new IllegalArgumentException("contentTypeId");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists(guid'" + listId + "')/ContentTypes('" + Util.encodeUrl(contentTypeId) + "')");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getListContentType", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ContentTypeHandler handler = new ServiceResponseUtil.ContentTypeHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getContentType();
    }

    /**
     * Gets the list content types.
     *
     * @param listId the list id
     * @return the list content types
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    private List<ContentType> getListContentTypes(String siteUrl, String listId) throws ServiceException {
        return getListContentTypes(siteUrl, listId, null);
    }

    /**
     * Gets the list content types.
     *
     * @param listId       the list id
     * @param queryOptions the query options
     * @return the list content types
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<ContentType> getListContentTypes(String siteUrl, String listId, List<IQueryOption> queryOptions) throws ServiceException {
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/ContentTypes");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getListContentTypes", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ContentTypesHandler handler = new ServiceResponseUtil.ContentTypesHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getContentTypes();
    }

    /**
     * Gets the list default view.
     *
     * @param listId the list id
     * @return the list default view
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public View getListDefaultView(String siteUrl, String listId) throws ServiceException {
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/DefaultView");
        if (callback.isDebug()) {
            callback.printDebug("getListDefaultView", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ListViewHandler handler = new ServiceResponseUtil.ListViewHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getListView();
    }

    /**
     * Gets the list event receivers.
     *
     * @param listId the list id
     * @return the list event receivers
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<EventReceiver> getListEventReceivers(String siteUrl, String listId) throws ServiceException {
        return getListEventReceivers(siteUrl, listId, null);
    }

    /**
     * Gets the list event receivers.
     *
     * @param listId       the list id
     * @param queryOptions the query options
     * @return the list event receivers
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<EventReceiver> getListEventReceivers(String siteUrl, String listId, List<IQueryOption> queryOptions) throws ServiceException {
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/EventReceivers");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getListEventReceivers", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.EventReceiversHandler handler = new ServiceResponseUtil.EventReceiversHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getEventReceivers();
    }

    /**
     * Gets the list schema xml.
     *
     * @param listId the list id
     * @return the list schema xml
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public String getListSchemaXml(String siteUrl, String listId) throws ServiceException {
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/SchemaXml");
        if (callback.isDebug()) {
            callback.printDebug("getListSchemaXml", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ListSchemaXmlHandler handler = new ServiceResponseUtil.ListSchemaXmlHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getListSchemaXml();
    }

    /**
     * Gets the list server settings.
     *
     * @param listId the list id
     * @return the list server settings
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public ListServerSettings getListServerSettings(String siteUrl, String listId) throws ServiceException, XMLStreamException, ParseException, IOException {
        // return new ListServerSettings(getListSchemaXml(siteUrl, listId));
        ListServerSettings serverSettings = new ListServerSettings();
        // serverSettings.setAttributeFromXml(getListSchemaXml(siteUrl, listId));
        try {
            serverSettings.setAttribute(getListSchemaXml(siteUrl, listId));
        } catch (Exception e) {
            throw new ServiceException(e.getMessage(), e);
        }
        return serverSettings;
    }

    /**
     * Gets the list regional settings.
     *
     * @param listId the list id
     * @return the list regional settings
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public ListRegionalSettings getListRegionalSettings(String siteUrl, String listId) throws ServiceException, XMLStreamException, ParseException, IOException {
        // return new RegionalSettings(getListSchemaXml(siteUrl, listId));
        ListRegionalSettings regionalSettings = new ListRegionalSettings();
        // regionalSettings.setAttributeFromXml(getListSchemaXml(siteUrl, listId));
        try {
            regionalSettings.setAttribute(getListSchemaXml(siteUrl, listId));
        } catch (Exception e) {
            throw new ServiceException(e.getMessage(), e);
        }
        return regionalSettings;
    }

    /**
     * Delete list.
     *
     * @param listId the list id
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public void deleteList(String siteUrl, String listId) throws ServiceException {
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')");
        if (callback.isDebug()) {
            callback.printDebug("deleteList", siteUrl, requestUrl.toString());
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), null, "DELETE", "*", null);
    }

    /**
     * Gets the content type field links.
     *
     * @param contentTypeId the content type id
     * @return the content type field links
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<FieldLink> getContentTypeFieldLinks(String siteUrl, String contentTypeId) throws ServiceException {
        return getContentTypeFieldLinks(siteUrl, contentTypeId, null);
    }

    /**
     * Gets the content type field links.
     *
     * @param contentTypeId the content type id
     * @param queryOptions  the query options
     * @return the content type field links
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<FieldLink> getContentTypeFieldLinks(String siteUrl, String contentTypeId, List<IQueryOption> queryOptions) throws ServiceException {
        if (contentTypeId == null) {
            throw new IllegalArgumentException("contentTypeId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/ContentTypes('" + Util.encodeUrl(contentTypeId) + "')/FieldLinks");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getContentTypeFieldLinks", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.FieldLinksHandler handler = new ServiceResponseUtil.FieldLinksHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getFieldLinks();
    }

    /**
     * Gets the content type fields.
     *
     * @param contentTypeId the content type id
     * @return the content type fields
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<Field> getContentTypeFields(String siteUrl, String contentTypeId) throws ServiceException {
        return getContentTypeFields(siteUrl, contentTypeId, null);
    }

    /**
     * Gets the content type fields.
     *
     * @param contentTypeId the content type id
     * @param queryOptions  the query options
     * @return the content type fields
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<Field> getContentTypeFields(String siteUrl, String contentTypeId, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (contentTypeId == null) {
            throw new IllegalArgumentException("contentTypeId");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/ContentTypes('" + Util.encodeUrl(contentTypeId) + "')/Fields");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getContentTypeFields", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.FieldsHandler handler = new ServiceResponseUtil.FieldsHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getFields();
    }

    /**
     * Gets the content type.
     *
     * @param contentTypeId the content type id
     * @return the content type
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    private ContentType getContentType(String siteUrl, String contentTypeId) throws ServiceException {
        return getContentType(siteUrl, contentTypeId, null);
    }

    public ContentType getContentType(String siteUrl, String contentTypeId, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (contentTypeId == null) {
            throw new IllegalArgumentException("contentTypeId");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/ContentTypes('" + Util.encodeUrl(contentTypeId) + "')");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getContentType", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ContentTypeHandler handler = new ServiceResponseUtil.ContentTypeHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getContentType();
    }

    /**
     * Gets the content types.
     *
     * @return the content types
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    private List<ContentType> getContentTypes(String siteUrl) throws ServiceException {
        return getContentTypes(siteUrl, null);
    }

    /**
     * Gets the content types.
     *
     * @param queryOptions the query options
     * @return the content types
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<ContentType> getContentTypes(String siteUrl, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/ContentTypes");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getContentTypes", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ContentTypesHandler handler = new ServiceResponseUtil.ContentTypesHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getContentTypes();
    }

    /**
     * Gets the associated member group.
     *
     * @return the associated member group
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public Group getAssociatedMemberGroup(String siteUrl) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/AssociatedMemberGroup");
        if (callback.isDebug()) {
            callback.printDebug("getAssociatedMemberGroup", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.GroupHandler handler = new ServiceResponseUtil.GroupHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getGroup();
    }

    /**
     * Gets the associated owner group.
     *
     * @return the associated owner group
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public Group getAssociatedOwnerGroup(String siteUrl) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/AssociatedOwnerGroup");
        if (callback.isDebug()) {
            callback.printDebug("getAssociatedOwnerGroup", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.GroupHandler handler = new ServiceResponseUtil.GroupHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getGroup();
    }

    /**
     * Gets the associated visitor group.
     *
     * @return the associated visitor group
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public Group getAssociatedVisitorGroup(String siteUrl) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/AssociatedVisitorGroup");
        if (callback.isDebug()) {
            callback.printDebug("getAssociatedVisitorGroup", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.GroupHandler handler = new ServiceResponseUtil.GroupHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getGroup();
    }

    /**
     * Gets the available content types.
     *
     * @return the available content types
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    private List<ContentType> getAvailableContentTypes(String siteUrl) throws ServiceException {
        return getAvailableContentTypes(siteUrl, null);
    }

    /**
     * Gets the available content types.
     *
     * @param queryOptions the query options
     * @return the available content types
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<ContentType> getAvailableContentTypes(String siteUrl, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/AvailableContentTypes");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getAvailableContentTypes", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ContentTypesHandler handler = new ServiceResponseUtil.ContentTypesHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getContentTypes();
    }

    /**
     * Gets the current user.
     *
     * @return the current user
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public User getCurrentUser(String siteUrl) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/CurrentUser");
        if (callback.isDebug()) {
            callback.printDebug("getCurrentUser", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.UserHandler handler = new ServiceResponseUtil.UserHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getUser();
    }

    /**
     * Gets the lists.
     *
     * @return the lists
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    private List<com.independentsoft.share.List> getLists(String siteUrl) throws ServiceException {
        return getLists(siteUrl, null);
    }

    /**
     * Gets the lists.
     *
     * @param queryOptions the query options
     * @return the lists
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<com.independentsoft.share.List> getLists(String siteUrl, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getLists", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ListsHandler handler = new ServiceResponseUtil.ListsHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getLists();
    }

    /**
     * Creates the view.
     *
     * @param listId the list id
     * @param view   the view
     * @return the view
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public View createView(String siteUrl, String listId, View view) throws ServiceException {
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (view == null) {
            throw new IllegalArgumentException("view");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/Lists('" + listId + "')/Views");
        String requestBody;
        try {
            requestBody = view.toCreateJSon(view.getProcessedListViewXml());
        } catch (Exception e) {
            throw new ServiceException(e.getMessage(), e, requestUrl.toString());
        }
        if (callback.isDebug()) {
            callback.printDebug("createView", siteUrl, requestUrl.toString(), requestBody);
        }
        ServiceResponseUtil.ListViewHandler handler = new ServiceResponseUtil.ListViewHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, handler);
        return handler.getListView();
    }

    /**
     * Update view.
     *
     * @param listId the list id
     * @param view   the view
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public void updateView(String siteUrl, String listId, View view) throws ServiceException {
        if (view == null) {
            throw new IllegalArgumentException("view");
        }
        try {
            updateView(siteUrl, listId, view.getId(), view.toUpdateJSon(view.getProcessedListViewXml()));
        } catch (Exception e) {
            throw new ServiceException(e.getMessage(), e);
        }
    }

    public void updateView(String siteUrl, String listId, String viewId, String requestBody) throws ServiceException {
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (viewId == null) {
            throw new IllegalArgumentException("viewId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/Lists('" + listId + "')/Views('" + viewId + "')");
        if (callback.isDebug()) {
            callback.printDebug("updateView", siteUrl, requestUrl.toString(), requestBody);
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, "MERGE", null);
    }

    /**
     * Adds the list role assignment.
     *
     * @param principalId the principal id
     * @param roleId      the role id
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean addListRoleAssignment(String siteUrl, String listId, int principalId, int roleId) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/roleassignments/addroleassignment(principalid=" + principalId + ",roledefid=" + roleId + ")");
        if (callback.isDebug()) {
            callback.printDebug("addRoleAssignment", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("AddRoleAssignment");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.isSuccess();
    }

    /**
     * Gets the views.
     *
     * @param listId the list id
     * @return the views
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<View> getViews(String siteUrl, String listId) throws ServiceException {
        return getViews(siteUrl, listId, null);
    }

    /**
     * Gets the views.
     *
     * @param listId       the list id
     * @param queryOptions the query options
     * @return the views
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<View> getViews(String siteUrl, String listId, List<IQueryOption> queryOptions) throws ServiceException {
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/Lists('" + listId + "')/Views");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getViews", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ListViewsHandler handler = new ServiceResponseUtil.ListViewsHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getListViews();
    }

    /**
     * Gets the view.
     *
     * @param listId the list id
     * @param viewId the view id
     * @return the view
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public View getView(String siteUrl, String listId, String viewId) throws ServiceException {
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (viewId == null) {
            throw new IllegalArgumentException("viewId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/Lists('" + listId + "')/Views('" + viewId + "')");
        if (callback.isDebug()) {
            callback.printDebug("getView", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ListViewHandler handler = new ServiceResponseUtil.ListViewHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getListView();
    }

    /**
     * Gets the view by title.
     *
     * @param listId the list id
     * @param title  the title
     * @return the view by title
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public View getViewByTitle(String siteUrl, String listId, String title) throws ServiceException {
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (title == null) {
            throw new IllegalArgumentException("title");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/Lists('" + listId + "')/Views/GetByTitle('" + Util.encodeUrl(title) + "')");
        if (callback.isDebug()) {
            callback.printDebug("getViewByTitle", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ListViewHandler handler = new ServiceResponseUtil.ListViewHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getListView();
    }

    /**
     * Gets the view html.
     *
     * @param listId the list id
     * @param viewId the view id
     * @return the view html
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public String getViewHtml(String siteUrl, String listId, String viewId) throws ServiceException {
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (viewId == null) {
            throw new IllegalArgumentException("viewId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/Lists('" + listId + "')/Views('" + viewId + "')/RenderAsHtml");
        if (callback.isDebug()) {
            callback.printDebug("getViewHtml", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.StringHandler handler = new ServiceResponseUtil.StringHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getString();
    }

    /**
     * Delete view.
     *
     * @param listId the list id
     * @param viewId the view id
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public void deleteView(String siteUrl, String listId, String viewId) throws ServiceException {
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (viewId == null) {
            throw new IllegalArgumentException("viewId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/Lists('" + listId + "')/Views('" + viewId + "')");
        if (callback.isDebug()) {
            callback.printDebug("deleteView", siteUrl, requestUrl.toString());
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), null, "DELETE", null);
    }

    /**
     * Gets the view fields.
     *
     * @param listId the list id
     * @param viewId the view id
     * @return the view fields
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<String> getViewFields(String siteUrl, String listId, String viewId) throws ServiceException {
        return getViewFields(siteUrl, listId, viewId, null);
    }

    /**
     * Gets the view fields.
     *
     * @param listId       the list id
     * @param viewId       the view id
     * @param queryOptions the query options
     * @return the view fields
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<String> getViewFields(String siteUrl, String listId, String viewId, List<IQueryOption> queryOptions) throws ServiceException {
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (viewId == null) {
            throw new IllegalArgumentException("viewId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/Lists('" + listId + "')/Views('" + viewId + "')/ViewFields");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getViewFields", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ListViewFieldsHandler handler = new ServiceResponseUtil.ListViewFieldsHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getListViewFields();
    }

    /**
     * Gets the view fields schema xml.
     *
     * @param listId the list id
     * @param viewId the view id
     * @return the view fields schema xml
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public String getViewFieldsSchemaXml(String siteUrl, String listId, String viewId) throws ServiceException {
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (viewId == null) {
            throw new IllegalArgumentException("viewId");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/Lists('" + listId + "')/Views('" + viewId + "')/ViewFields");
        if (callback.isDebug()) {
            callback.printDebug("getViewFieldsSchemaXml", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ListSchemaXmlHandler handler = new ServiceResponseUtil.ListSchemaXmlHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getListSchemaXml();
    }

    /**
     * Creates the view field.
     *
     * @param listId    the list id
     * @param viewId    the view id
     * @param fieldName the field name
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean createViewField(String siteUrl, String listId, String viewId, String fieldName) throws ServiceException {
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (viewId == null) {
            throw new IllegalArgumentException("viewId");
        }
        if (fieldName == null) {
            throw new IllegalArgumentException("fieldName");
        }
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/Lists('" + listId + "')/Views('" + viewId + "')/ViewFields/AddViewField('" + Util.encodeUrl(fieldName) + "')");
        if (callback.isDebug()) {
            callback.printDebug("createViewField", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("AddViewField");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.isSuccess();
    }

    /**
     * Delete view field.
     *
     * @param listId    the list id
     * @param viewId    the view id
     * @param fieldName the field name
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean deleteViewField(String siteUrl, String listId, String viewId, String fieldName) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (viewId == null) {
            throw new IllegalArgumentException("viewId");
        }
        if (fieldName == null) {
            throw new IllegalArgumentException("fieldName");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/Lists('" + listId + "')/Views('" + viewId + "')/ViewFields/RemoveViewField('" + Util.encodeUrl(fieldName) + "')");
        if (callback.isDebug()) {
            callback.printDebug("deleteViewField", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("RemoveViewField");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.isSuccess();
    }

    /**
     * Delete all view fields.
     *
     * @param listId the list id
     * @param viewId the view id
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean deleteAllViewFields(String siteUrl, String listId, String viewId) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (viewId == null) {
            throw new IllegalArgumentException("viewId");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/Lists('" + listId + "')/Views('" + viewId + "')/ViewFields/RemoveAllViewFields");
        if (callback.isDebug()) {
            callback.printDebug("deleteAllViewFields", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("RemoveAllViewFields");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.isSuccess();
    }

    /**
     * Move view field.
     *
     * @param listId    the list id
     * @param viewId    the view id
     * @param fieldName the field name
     * @param index     the index
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean moveViewField(String siteUrl, String listId, String viewId, String fieldName, int index) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (viewId == null) {
            throw new IllegalArgumentException("viewId");
        }
        if (fieldName == null) {
            throw new IllegalArgumentException("fieldName");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/Lists('" + listId + "')/Views('" + viewId + "')/ViewFields/MoveViewFieldTo");
        String requestBody = "{ 'field': '" + Util.encodeJson(fieldName) + "', 'index': " + index + " }";
        if (callback.isDebug()) {
            callback.printDebug("moveViewField", siteUrl, requestUrl.toString(), requestBody);
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("MoveViewFieldTo");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, handler);
        return handler.isSuccess();
    }

    /**
     * Gets the event receivers.
     *
     * @return the event receivers
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<EventReceiver> getEventReceivers(String siteUrl) throws ServiceException {
        return getEventReceivers(siteUrl, null);
    }

    /**
     * Gets the event receivers.
     *
     * @param queryOptions the query options
     * @return the event receivers
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<EventReceiver> getEventReceivers(String siteUrl, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/EventReceivers");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getEventReceivers", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.EventReceiversHandler handler = new ServiceResponseUtil.EventReceiversHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getEventReceivers();
    }

    /**
     * Gets the event receiver.
     *
     * @param receiverId the receiver id
     * @return the event receiver
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public EventReceiver getEventReceiver(String siteUrl, String receiverId) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (receiverId == null) {
            throw new IllegalArgumentException("receiverId");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/EventReceivers('" + receiverId + "')");
        if (callback.isDebug()) {
            callback.printDebug("getEventReceiver", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.EventReceiverHandler handler = new ServiceResponseUtil.EventReceiverHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getEventReceiver();
    }

    /**
     * Gets the list templates.
     *
     * @return the list templates
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<ListTemplate> getListTemplates(String siteUrl) throws ServiceException {
        return getListTemplates(siteUrl, null);
    }

    /**
     * Gets the list templates.
     *
     * @param queryOptions the query options
     * @return the list templates
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<ListTemplate> getListTemplates(String siteUrl, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/ListTemplates");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getListTemplates", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ListTemplatesHandler handler = new ServiceResponseUtil.ListTemplatesHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getListTemplates();
    }

    /**
     * Gets the list template.
     *
     * @param name the name
     * @return the list template
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public ListTemplate getListTemplate(String siteUrl, String name) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/ListTemplates('" + Util.encodeUrl(name) + "')");
        if (callback.isDebug()) {
            callback.printDebug("getListTemplate", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ListTemplateHandler handler = new ServiceResponseUtil.ListTemplateHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getListTemplate();
    }

    /**
     * Gets the custom list templates.
     *
     * @return the custom list templates
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<ListTemplate> getCustomListTemplates(String siteUrl) throws ServiceException {
        return getCustomListTemplates(siteUrl, null);
    }

    /**
     * Gets the custom list templates.
     *
     * @param queryOptions the query options
     * @return the custom list templates
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<ListTemplate> getCustomListTemplates(String siteUrl, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/GetCustomListTemplates");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getCustomListTemplates", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ListTemplatesHandler handler = new ServiceResponseUtil.ListTemplatesHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getListTemplates();
    }

    /**
     * Gets the file stream.
     *
     * @param filePath the file path
     * @return the file stream
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public InputStream getFileStream(String siteUrl, String filePath) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (filePath == null) {
            throw new IllegalArgumentException("filePath");
        }

        // 22245: Support % and # in files and folders, handle ' in filePath
        // StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativeUrl(@v)/$value?@v='" + Util.encodeUrl(filePath) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativePath(decodedUrl=@v)/$value?@v='" + Util.encodeUrl(Util.escapeQueryUrl(filePath)) + "'");
        if (callback.isDebug()) {
            callback.printDebug("getFileStream", siteUrl, requestUrl.toString());
        }
        try {
            return doSendRawRequest(siteUrl, "GET", requestUrl.toString(), null, null, null, null, false, false);
        } catch (ServiceException e) {
            throw e;
        } catch (Exception e) {
            throw new ServiceException(e.getMessage(), e, requestUrl.toString());
        }
    }

    /**
     * Gets the input stream.
     *
     * @param url the url
     * @return the input stream
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public InputStream getInputStream(String siteUrl, String url) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (url == null) {
            throw new IllegalArgumentException("url");
        }

        String requestUrl = Util.encodeUrlInputStream(url);
        if (callback.isDebug()) {
            callback.printDebug("getInputStream", siteUrl, requestUrl);
        }
        try {
            return doSendRawRequest(siteUrl, "GET", requestUrl, null, null, null, null, false, false);
        } catch (ServiceException e) {
            throw e;
        } catch (Exception e) {
            throw new ServiceException(e.getMessage(), e, requestUrl);
        }
    }

    /**
     * Gets the file content.
     *
     * @param filePath the file path
     * @return the file content
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public byte[] getFileContent(String siteUrl, String filePath) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (filePath == null) {
            throw new IllegalArgumentException("filePath");
        }

        // 22245: Support % and # in files and folders
        // StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativeUrl(@v)/$value?@v='" + Util.encodeUrl(filePath) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativePath(decodedUrl=@v)/$value?@v='" + Util.encodeUrl(Util.escapeQueryUrl(filePath)) + "'");
        if (callback.isDebug()) {
            callback.printDebug("getFileContent", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.FileContentHandler handler = new ServiceResponseUtil.FileContentHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getFileContent();
    }

    /**
     * Delete file.
     *
     * @param filePath the file path
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public void deleteFile(String siteUrl, String filePath) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (filePath == null) {
            throw new IllegalArgumentException("filePath");
        }

        // 22245: Support % and # in files and folders
        // StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativeUrl(@v)?@v='" + Util.encodeUrl(filePath) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativePath(decodedUrl=@v)?@v='" + Util.encodeUrl(Util.escapeQueryUrl(filePath)) + "'");
        if (callback.isDebug()) {
            callback.printDebug("deleteFile", siteUrl, requestUrl.toString());
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), null, "DELETE", null);
    }

    /**
     * Creates the file.
     *
     * @param filePath the file path
     * @param buffer   the buffer
     * @return the file
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public File createFile(String siteUrl, String filePath, byte[] buffer) throws ServiceException {
        return createFile(siteUrl, filePath, buffer, false);
    }

    /**
     * Creates the file.
     *
     * @param filePath  the file path
     * @param buffer    the buffer
     * @param overwrite the overwrite
     * @return the file
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public File createFile(String siteUrl, String filePath, byte[] buffer, boolean overwrite) throws ServiceException {
        ByteArrayInputStream stream = new ByteArrayInputStream(buffer);
        return createFile(siteUrl, filePath, stream, overwrite);
    }

    /**
     * Creates the file.
     *
     * @param filePath the file path
     * @param stream   the stream
     * @return the file
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public File createFile(String siteUrl, String filePath, InputStream stream) throws ServiceException {
        return createFile(siteUrl, filePath, stream, false);
    }

    /**
     * Creates the file.
     *
     * @param filePath  the file path
     * @param stream    the stream
     * @param overwrite the overwrite
     * @return the file
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    // [Start] 23474: Support query option
    public File createFile(String siteUrl, String filePath, InputStream stream, boolean overwrite) throws ServiceException {
        return createFile(siteUrl, filePath, stream, overwrite, null);
    }

    public File createFile(String siteUrl, String filePath, InputStream stream, boolean overwrite, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (filePath == null) {
            throw new IllegalArgumentException("filePath");
        }
        if (stream == null) {
            throw new IllegalArgumentException("stream");
        }

        int index = filePath.lastIndexOf("/");
        String folderPath = index > 0 ? filePath.substring(0, index) : "/";
        String fileName = filePath.substring(index + 1);
        // 22245: Support % and # in files and folders
        // 22692: Use decodedUrl as folder path to handle special chars
        // String requestUrl = "_api/web/GetFolderByServerRelativeUrl(@v)/files/Add(overwrite=" + Boolean.toString(overwrite).toLowerCase() + ",url='" + Util.encodeUrl(fileName) + "')?@v='" + Util.encodeUrl(folderPath) + "'");
        // String requestUrl = "_api/web/GetFolderByServerRelativePath(decodedUrl=@v)/files/Add(overwrite=" + Boolean.toString(overwrite).toLowerCase() + ",url='" + Util.encodeUrl(fileName) + "')?@v='" + Util.encodeUrl(folderPath) + "'");
        // String requestUrl = "_api/web/GetFolderByServerRelativePath(decodedUrl=@v)/files/AddUsingPath(overwrite=" + Boolean.toString(overwrite).toLowerCase() + ",decodedUrl='" + Util.encodeUrl(Util.escapeQueryUrl(fileName)) + "')?@v='" + Util.encodeUrl(Util.escapeQueryUrl(folderPath)) + "'");
        // String queryOptionsString = Util.queryOptionsToString(queryOptions, true);
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions, true);
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFolderByServerRelativePath(decodedUrl=@v)/files/AddUsingPath(overwrite=" + Boolean.toString(overwrite).toLowerCase() + ",decodedUrl='" + Util.encodeUrl(Util.escapeQueryUrl(fileName)) + "')?@v='" + Util.encodeUrl(Util.escapeQueryUrl(folderPath)) + "'");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("createFile", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.FileHandler handler = new ServiceResponseUtil.FileHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), null, null, null, stream, true, false, handler);
        return handler.getFile();
    }
    // [End] 23474

    /**
     * Creates the template file.
     *
     * @param filePath the file path
     * @return the file
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public File createTemplateFile(String siteUrl, String filePath) throws ServiceException {
        return createTemplateFile(siteUrl, filePath, null);
    }

    /**
     * Creates the template file.
     *
     * @param filePath the file path
     * @param type     the type
     * @return the file
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public File createTemplateFile(String siteUrl, String filePath, TemplateFileType type) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (filePath == null) {
            throw new IllegalArgumentException("filePath");
        }

        int index = filePath.lastIndexOf("/");
        String folderPath = index > 0 ? filePath.substring(0, index) : "/";
        String templateFileTypeString = (type != null) ? ",templatefiletype=" + type.getValue() : "";
        // 22245: Support % and # in files and folders
        // String requestUrl = new StringBuilder("_api/web/GetFolderByServerRelativeUrl(@v)/files/AddTemplateFile(urloffile='" + Util.encodeUrl(filePath) + "'" + templateFileTypeString + ")?@v='" + Util.encodeUrl(folderPath) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFolderByServerRelativePath(decodedUrl=@v)/files/AddTemplateFile(urloffile='");
        requestUrl.append(Util.encodeUrl(Util.escapeQueryUrl(filePath)));
        requestUrl.append("'");
        requestUrl.append(templateFileTypeString);
        requestUrl.append(")?@v='");
        requestUrl.append(Util.encodeUrl(Util.escapeQueryUrl(folderPath)));
        requestUrl.append("'");
        if (callback.isDebug()) {
            callback.printDebug("createTemplateFile", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.FileHandler handler = new ServiceResponseUtil.FileHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.getFile();
    }

    /**
     * Update file content.
     *
     * @param filePath the file path
     * @param buffer   the buffer
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public void updateFileContent(String siteUrl, String filePath, byte[] buffer) throws ServiceException {
        ByteArrayInputStream stream = new ByteArrayInputStream(buffer);
        updateFileContent(siteUrl, filePath, stream);
    }

    /**
     * Update file content.
     *
     * @param filePath the file path
     * @param stream   the stream
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public void updateFileContent(String siteUrl, String filePath, InputStream stream) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (filePath == null) {
            throw new IllegalArgumentException("filePath");
        }
        if (stream == null) {
            throw new IllegalArgumentException("stream");
        }

        // 22245: Support % and # in files and folders
        // String requestUrl = new StringBuilder("_api/web/GetFileByServerRelativeUrl(@v)/$value?@v='" + Util.encodeUrl(filePath) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativePath(decodedUrl=@v)/$value?@v='");
        requestUrl.append(Util.encodeUrl(Util.escapeQueryUrl(filePath)));
        requestUrl.append("'");
        if (callback.isDebug()) {
            callback.printDebug("updateFileContent", siteUrl, requestUrl.toString());
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), null, "PUT", null, stream, true, false, null);
    }

    public enum UpdateFileMode {
        Start, Continue, Finish, Cancel;
    }

    public void updateFileContent(String siteUrl,
                                  String filePath,
                                  InputStream stream,
                                  UpdateFileMode mode,
                                  String guid,
                                  long offset)
            throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (filePath == null) {
            throw new IllegalArgumentException("filePath");
        }
        if (stream == null) {
            throw new IllegalArgumentException("stream");
        }
        if (guid == null || "".equals(guid)) {
            throw new IllegalArgumentException("guid");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativePath(decodedUrl=@v)");
        if (mode == UpdateFileMode.Start) {
            requestUrl.append("/startupload(uploadId=guid'" + guid + "')");
        } else if (mode == UpdateFileMode.Continue) {
            requestUrl.append("/continueupload(uploadId=guid'" + guid + "',fileOffset=" + offset + ")");
        } else if (mode == UpdateFileMode.Finish) {
            requestUrl.append("/finishupload(uploadId=guid'" + guid + "',fileOffset=" + offset + ")");
        } else {
            requestUrl.append("/cancelupload(uploadId=guid'" + guid + "')");
        }
        requestUrl.append("?@v='" + Util.encodeUrl(Util.escapeQueryUrl(filePath)) + "'");
        if (callback.isDebug()) {
            callback.printDebug("updateFileContent", siteUrl, requestUrl.toString());
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), null, null, null, stream, true, false, null);
    }

    /**
     * Gets the file.
     *
     * @param filePath the file path
     * @return the file
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public File getFile(String siteUrl, String filePath) throws ServiceException {
        return getFile(siteUrl, filePath, null);
    }

    public File getFile(String siteUrl, String filePath, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (filePath == null) {
            throw new IllegalArgumentException("filePath");
        }

        // 22245: Support % and # in files and folders
        // String requestUrl = new StringBuilder("_api/web/GetFileByServerRelativeUrl(@v)?@v='" + Util.encodeUrl(filePath) + "'");
        //String queryOptionsString = queryOptions != null ? Util.queryOptionsToString(queryOptions, true) : "");

        StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativePath(decodedUrl=@v)?@v='" + Util.encodeUrl(Util.escapeQueryUrl(filePath)) + "'");
        if (queryOptions != null) {
            StringBuilder sbQuery = new StringBuilder("");
            Util.queryOptionsToString(sbQuery, queryOptions, true);
            requestUrl.append(sbQuery);
        }
        if (callback.isDebug()) {
            callback.printDebug("getFile", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.FileHandler handler = new ServiceResponseUtil.FileHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getFile();
    }

    /**
     * Recycle file.
     *
     * @param filePath the file path
     * @return the string
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public String recycleFile(String siteUrl, String filePath) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (filePath == null) {
            throw new IllegalArgumentException("filePath");
        }

        // 22245: Support % and # in files and folders
        // String requestUrl = new StringBuilder("_api/web/GetFileByServerRelativeUrl(@v)/recycle?@v='" + Util.encodeUrl(filePath) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativePath(decodedUrl=@v)/recycle?@v='" + Util.encodeUrl(Util.escapeQueryUrl(filePath)) + "'");
        if (callback.isDebug()) {
            callback.printDebug("recycleFile", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.RecycleHandler handler = new ServiceResponseUtil.RecycleHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.getRecycle();
    }

    /**
     * Recycle list.
     *
     * @param listId the list id
     * @return the string
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public String recycleList(String siteUrl, String listId) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/Recycle");
        if (callback.isDebug()) {
            callback.printDebug("recycleList", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.RecycleHandler handler = new ServiceResponseUtil.RecycleHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.getRecycle();
    }

    /**
     * Recycle list item.
     *
     * @param listId the list id
     * @param itemId the item id
     * @return the string
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public String recycleListItem(String siteUrl, String listId, int itemId) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (itemId < 0) {
            throw new IllegalArgumentException("The parameter itemId must be non-negative.");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/items(" + itemId + ")/Recycle");
        if (callback.isDebug()) {
            callback.printDebug("recycleListItem", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.RecycleHandler handler = new ServiceResponseUtil.RecycleHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.getRecycle();
    }

    /**
     * Unpublish.
     *
     * @param filePath the file path
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean unpublish(String siteUrl, String filePath) throws ServiceException {
        return unpublish(siteUrl, filePath, null);
    }

    /**
     * Unpublish.
     *
     * @param filePath the file path
     * @param comment  the comment
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean unpublish(String siteUrl, String filePath, String comment) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (filePath == null) {
            throw new IllegalArgumentException("filePath");
        }

        // 22245: Support % and # in files and folders
        // String requestUrl = new StringBuilder("_api/web/GetFileByServerRelativeUrl(@v)/Unpublish?@v='" + Util.encodeUrl(filePath) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativePath(decodedUrl=@v)/Unpublish?@v='" + Util.encodeUrl(Util.escapeQueryUrl(filePath)) + "'");
        if (comment != null && comment.length() > 0) {
            requestUrl.append("('" + Util.encodeUrl(comment) + "')");
        }
        if (callback.isDebug()) {
            callback.printDebug("unpublish", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("UnPublish");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.isSuccess();
    }

    /**
     * Publish.
     *
     * @param filePath the file path
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean publish(String siteUrl, String filePath) throws ServiceException {
        return publish(siteUrl, filePath, null);
    }

    /**
     * Publish.
     *
     * @param filePath the file path
     * @param comment  the comment
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean publish(String siteUrl, String filePath, String comment) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (filePath == null) {
            throw new IllegalArgumentException("filePath");
        }

        // 22245: Support % and # in files and folders
        // String requestUrl = new StringBuilder("_api/web/GetFileByServerRelativeUrl(@v)/Publish?@v='" + Util.encodeUrl(filePath) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativePath(decodedUrl=@v)/Publish?@v='" + Util.encodeUrl(Util.escapeQueryUrl(filePath)) + "'");
        if (comment != null && comment.length() > 0) {
            requestUrl.append("('" + Util.encodeUrl(comment) + "')");
        }
        if (callback.isDebug()) {
            callback.printDebug("publish", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("Publish");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.isSuccess();
    }

    /**
     * Approve.
     *
     * @param filePath the file path
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean approve(String siteUrl, String filePath) throws ServiceException {
        return approve(siteUrl, filePath, null);
    }

    /**
     * Approve.
     *
     * @param filePath the file path
     * @param comment  the comment
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean approve(String siteUrl, String filePath, String comment) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (filePath == null) {
            throw new IllegalArgumentException("filePath");
        }

        // 22245: Support % and # in files and folders
        // String requestUrl = new StringBuilder("_api/web/GetFileByServerRelativeUrl(@v)/Approve?@v='" + Util.encodeUrl(filePath) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativePath(decodedUrl=@v)/Approve?@v='" + Util.encodeUrl(Util.escapeQueryUrl(filePath)) + "'");
        if (comment != null && comment.length() > 0) {
            requestUrl.append("('" + Util.encodeUrl(comment) + "')");
        }
        if (callback.isDebug()) {
            callback.printDebug("approve", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("Approve");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.isSuccess();
    }

    /**
     * Deny.
     *
     * @param filePath the file path
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean deny(String siteUrl, String filePath) throws ServiceException {
        return deny(siteUrl, filePath, null);
    }

    /**
     * Deny.
     *
     * @param filePath the file path
     * @param comment  the comment
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean deny(String siteUrl, String filePath, String comment) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (filePath == null) {
            throw new IllegalArgumentException("filePath");
        }

        // 22245: Support % and # in files and folders
        // String requestUrl = new StringBuilder("_api/web/GetFileByServerRelativeUrl(@v)/Deny?@v='" + Util.encodeUrl(filePath) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativePath(decodedUrl=@v)/Deny?@v='" + Util.encodeUrl(Util.escapeQueryUrl(filePath)) + "'");
        if (comment != null && comment.length() > 0) {
            requestUrl.append("('" + Util.encodeUrl(comment) + "')");
        }
        if (callback.isDebug()) {
            callback.printDebug("deny", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("Deny");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.isSuccess();
    }

    /**
     * Undo check out.
     *
     * @param filePath the file path
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean undoCheckOut(String siteUrl, String filePath) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (filePath == null) {
            throw new IllegalArgumentException("filePath");
        }

        // 22245: Support % and # in files and folders
        // String requestUrl = new StringBuilder("_api/web/GetFileByServerRelativeUrl(@v)/UndoCheckOut()?@v='" + Util.encodeUrl(filePath) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativePath(decodedUrl=@v)/UndoCheckOut()?@v='" + Util.encodeUrl(Util.escapeQueryUrl(filePath)) + "'");
        if (callback.isDebug()) {
            callback.printDebug("undoCheckOut", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("UndoCheckOut");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.isSuccess();
    }

    /**
     * Check out.
     *
     * @param filePath the file path
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean checkOut(String siteUrl, String filePath) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (filePath == null) {
            throw new IllegalArgumentException("filePath");
        }

        // 22245: Support % and # in files and folders
        // String requestUrl = new StringBuilder("_api/web/GetFileByServerRelativeUrl(@v)/Checkout()?@v='" + Util.encodeUrl(filePath) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativePath(decodedUrl=@v)/Checkout()?@v='" + Util.encodeUrl(Util.escapeQueryUrl(filePath)) + "'");
        if (callback.isDebug()) {
            callback.printDebug("checkOut", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("CheckOut");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.isSuccess();
    }

    /**
     * Check in.
     *
     * @param filePath the file path
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean checkIn(String siteUrl, String filePath) throws ServiceException {
        return checkIn(siteUrl, filePath, CheckInType.getInstance(CheckInType.Defines.MINOR));
    }

    /**
     * Check in.
     *
     * @param filePath the file path
     * @param comment  the comment
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean checkIn(String siteUrl, String filePath, String comment) throws ServiceException {
        return checkIn(siteUrl, filePath, CheckInType.getInstance(CheckInType.Defines.MINOR), comment);
    }

    /**
     * Check in.
     *
     * @param filePath the file path
     * @param type     the type
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean checkIn(String siteUrl, String filePath, CheckInType type) throws ServiceException {
        return checkIn(siteUrl, filePath, type, null);
    }

    /**
     * Check in.
     *
     * @param filePath the file path
     * @param type     the type
     * @param comment  the comment
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean checkIn(String siteUrl, String filePath, CheckInType type, String comment) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (filePath == null) {
            throw new IllegalArgumentException("filePath");
        }

        // 22245: Support % and # in files and folders
        // String requestUrl = new StringBuilder("_api/web/GetFileByServerRelativeUrl(@v)/Checkin?@v='" + Util.encodeUrl(filePath) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativePath(decodedUrl=@v)/Checkin");
        String encodedComment = comment != null ? "comment='" + Util.encodeUrl(comment) + "'," : "";
        requestUrl.append("(" + encodedComment + "checkintype=" + type.getValue() + ")");
        requestUrl.append("?@v='" + Util.encodeUrl(Util.escapeQueryUrl(filePath)) + "'");
        if (callback.isDebug()) {
            callback.printDebug("checkIn", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("CheckIn");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.isSuccess();
    }

    /**
     * Copy file.
     *
     * @param sourceFilePath      the source file path
     * @param destinationFilePath the destination file path
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean copyFile(String siteUrl, String sourceFilePath, String destinationFilePath) throws ServiceException {
        return copyFile(siteUrl, sourceFilePath, destinationFilePath, false);
    }

    /**
     * Copy file.
     *
     * @param sourceFilePath      the source file path
     * @param destinationFilePath the destination file path
     * @param overwrite           the overwrite
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean copyFile(String siteUrl, String sourceFilePath, String destinationFilePath, boolean overwrite) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (sourceFilePath == null) {
            throw new IllegalArgumentException("sourceFilePath");
        }
        if (destinationFilePath == null) {
            throw new IllegalArgumentException("destinationFilePath");
        }

        // 22245: Support % and # in files and folders
        // String requestUrl = new StringBuilder("_api/web/GetFileByServerRelativeUrl(@v)/CopyTo?@v='" + Util.encodeUrl(sourceFilePath) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativePath(decodedUrl=@v)/CopyTo");
        requestUrl.append("(strnewurl='" + Util.encodeUrl(Util.escapeQueryUrl(destinationFilePath)) + "',boverwrite=" + Boolean.toString(overwrite).toLowerCase() + ")")
        ;
        requestUrl.append("?@v='" + Util.encodeUrl(Util.escapeQueryUrl(sourceFilePath)) + "'");
        if (callback.isDebug()) {
            callback.printDebug("copyFile", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("CopyTo");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.isSuccess();
    }

    /**
     * Move file.
     *
     * @param sourceFilePath      the source file path
     * @param destinationFilePath the destination file path
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean moveFile(String siteUrl, String sourceFilePath, String destinationFilePath) throws ServiceException {
        return moveFile(siteUrl, sourceFilePath, destinationFilePath, null);
    }

    /**
     * Move file.
     *
     * @param sourceFilePath      the source file path
     * @param destinationFilePath the destination file path
     * @param operation           the operation
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean moveFile(String siteUrl, String sourceFilePath, String destinationFilePath, MoveOperation operation) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (sourceFilePath == null) {
            throw new IllegalArgumentException("sourceFilePath");
        }
        if (destinationFilePath == null) {
            throw new IllegalArgumentException("destinationFilePath");
        }

        // 22245: Support % and # in files and folders
        // String requestUrl = new StringBuilder("_api/web/GetFileByServerRelativeUrl(@v)/MoveTo");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativePath(decodedUrl=@v)/MoveTo");
        String moveOperationString = operation != null ? ",flags=" + operation.getValue() : "";
        requestUrl.append("(newurl='" + Util.encodeUrl(Util.escapeQueryUrl(destinationFilePath)) + "'" + moveOperationString + ")")
        ;
        requestUrl.append("?@v='" + Util.encodeUrl(Util.escapeQueryUrl(sourceFilePath)) + "'");
        if (callback.isDebug()) {
            callback.printDebug("moveFile", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("MoveTo");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.isSuccess();
    }

    /**
     * Gets the limited web part manager.
     *
     * @param filePath the file path
     * @return the limited web part manager
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public LimitedWebPartManager getLimitedWebPartManager(String siteUrl, String filePath) throws ServiceException {
        return getLimitedWebPartManager(siteUrl, filePath, PersonalizationScope.getInstance(PersonalizationScope.Defines.USER));
    }

    /**
     * Gets the limited web part manager.
     *
     * @param filePath the file path
     * @param scope    the scope
     * @return the limited web part manager
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public LimitedWebPartManager getLimitedWebPartManager(String siteUrl, String filePath, PersonalizationScope scope) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (filePath == null) {
            throw new IllegalArgumentException("filePath");
        }

        // 22245: Support % and # in files and folders
        // String requestUrl = new StringBuilder("_api/web/GetFileByServerRelativeUrl(@v)/GetLimitedWebPartManager(scope=" + scope.getValue() + ")?@v='" + Util.encodeUrl(filePath) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativePath(decodedUrl=@v)/GetLimitedWebPartManager(scope=" + scope.getValue() + ")?@v='" + Util.encodeUrl(Util.escapeQueryUrl(filePath)) + "'");
        if (callback.isDebug()) {
            callback.printDebug("getLimitedWebPartManager", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.LimitedWebPartManagerHandler handler = new ServiceResponseUtil.LimitedWebPartManagerHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getLimitedWebPartManager();
    }

    /**
     * Gets the file author.
     *
     * @param filePath the file path
     * @return the file author
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public User getFileAuthor(String siteUrl, String filePath) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (filePath == null) {
            throw new IllegalArgumentException("filePath");
        }

        // 22245: Support % and # in files and folders
        // String requestUrl = new StringBuilder("_api/web/GetFileByServerRelativeUrl(@v)/Author?@v='" + Util.encodeUrl(filePath) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativePath(decodedUrl=@v)/Author?@v='" + Util.encodeUrl(Util.escapeQueryUrl(filePath)) + "'");
        if (callback.isDebug()) {
            callback.printDebug("getFileAuthor", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.UserHandler handler = new ServiceResponseUtil.UserHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getUser();
    }

    /**
     * Checked out by user.
     *
     * @param filePath the file path
     * @return the user
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public User checkedOutByUser(String siteUrl, String filePath) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (filePath == null) {
            throw new IllegalArgumentException("filePath");
        }

        // 22245: Support % and # in files and folders
        // String requestUrl = new StringBuilder("_api/web/GetFileByServerRelativeUrl(@v)/CheckedOutByUser?@v='" + Util.encodeUrl(filePath) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativePath(decodedUrl=@v)/CheckedOutByUser?@v='" + Util.encodeUrl(Util.escapeQueryUrl(filePath)) + "'");
        if (callback.isDebug()) {
            callback.printDebug("checkedOutByUser", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.UserHandler handler = new ServiceResponseUtil.UserHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getUser();
    }

    /**
     * Gets the modified by user.
     *
     * @param filePath the file path
     * @return the modified by user
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public User getModifiedByUser(String siteUrl, String filePath) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (filePath == null) {
            throw new IllegalArgumentException("filePath");
        }

        // 22245: Support % and # in files and folders
        // StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativeUrl(@v)/ModifiedBy?@v='" + Util.encodeUrl(filePath) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativePath(decodedUrl=@v)/ModifiedBy?@v='" + Util.encodeUrl(Util.escapeQueryUrl(filePath)) + "'");
        if (callback.isDebug()) {
            callback.printDebug("getModifiedByUser", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.UserHandler handler = new ServiceResponseUtil.UserHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getUser();
    }

    /**
     * Gets the locked by user.
     *
     * @param filePath the file path
     * @return the locked by user
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public User getLockedByUser(String siteUrl, String filePath) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (filePath == null) {
            throw new IllegalArgumentException("filePath");
        }

        // 22245: Support % and # in files and folders
        // StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativeUrl(@v)/LockedByUser?@v='" + Util.encodeUrl(filePath) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativePath(decodedUrl=@v)/LockedByUser?@v='" + Util.encodeUrl(Util.escapeQueryUrl(filePath)) + "'");
        if (callback.isDebug()) {
            callback.printDebug("getLockedByUser", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.UserHandler handler = new ServiceResponseUtil.UserHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getUser();
    }

    /**
     * Gets the created by user.
     *
     * @param filePath  the file path
     * @param versionId the version id
     * @return the created by user
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public User getCreatedByUser(String siteUrl, String filePath, int versionId) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (filePath == null) {
            throw new IllegalArgumentException("filePath");
        }

        // 22245: Support % and # in files and folders
        // StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativeUrl(@v)/Versions(" + Integer.toString(versionId) + ")/CreatedBy?@v='" + Util.encodeUrl(filePath) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativePath(decodedUrl=@v)/Versions(" + Integer.toString(versionId) + ")/CreatedBy?@v='" + Util.encodeUrl(Util.escapeQueryUrl(filePath)) + "'");
        if (callback.isDebug()) {
            callback.printDebug("getCreatedByUser", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.UserHandler handler = new ServiceResponseUtil.UserHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getUser();
    }

    /**
     * Gets the file versions.
     *
     * @param filePath the file path
     * @return the file versions
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<FileVersion> getFileVersions(String siteUrl, String filePath) throws ServiceException {
        return getFileVersions(siteUrl, filePath, null);
    }

    /**
     * Gets the file versions.
     *
     * @param filePath     the file path
     * @param queryOptions the query options
     * @return the file versions
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<FileVersion> getFileVersions(String siteUrl, String filePath, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (filePath == null) {
            throw new IllegalArgumentException("filePath");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        // queryOptionsString = queryOptionsString.isEmpty() ? queryOptionsString : "&" + queryOptionsString.substring(1);
        // 22245: Support % and # in files and folders
        // String requestUrl = "_api/web/GetFileByServerRelativeUrl(@v)/Versions?@v='" + Util.encodeUrl(filePath) + "'" + queryOptionsString;
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativePath(decodedUrl=@v)/Versions?@v='" + Util.encodeUrl(Util.escapeQueryUrl(filePath)) + "'");
        if (sbQuery.length() > 0) {
            requestUrl.append("&");
            requestUrl.append(sbQuery);
        }
        if (callback.isDebug()) {
            callback.printDebug("getFileVersions", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.FileVersionsHandler handler = new ServiceResponseUtil.FileVersionsHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getFileVersions();
    }

    /**
     * Gets the file version.
     *
     * @param filePath  the file path
     * @param versionId the version id
     * @return the file version
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public FileVersion getFileVersion(String siteUrl, String filePath, int versionId) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (filePath == null) {
            throw new IllegalArgumentException("filePath");
        }

        // 22245: Support % and # in files and folders
        // StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativeUrl(@v)/Versions(" + Integer.toString(versionId) + ")?@v='" + Util.encodeUrl(filePath) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativePath(decodedUrl=@v)/Versions(" + Integer.toString(versionId) + ")?@v='" + Util.encodeUrl(Util.escapeQueryUrl(filePath)) + "'");
        if (callback.isDebug()) {
            callback.printDebug("getFileVersion", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.FileVersionHandler handler = new ServiceResponseUtil.FileVersionHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getFileVersion();
    }

    /**
     * Delete file version.
     *
     * @param filePath  the file path
     * @param versionId the version id
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean deleteFileVersion(String siteUrl, String filePath, int versionId) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (filePath == null) {
            throw new IllegalArgumentException("filePath");
        }

        // 22245: Support % and # in files and folders
        // StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativeUrl(@v)/Versions/DeleteById(vid=" + Integer.toString(versionId) + ")?@v='" + Util.encodeUrl(filePath) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativePath(decodedUrl=@v)/Versions/DeleteById(vid=" + Integer.toString(versionId) + ")?@v='" + Util.encodeUrl(Util.escapeQueryUrl(filePath)) + "'");
        if (callback.isDebug()) {
            callback.printDebug("deleteFileVersion", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("DeleteByID");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.isSuccess();
    }

    /**
     * Delete all file versions.
     *
     * @param filePath the file path
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean deleteAllFileVersions(String siteUrl, String filePath) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (filePath == null) {
            throw new IllegalArgumentException("filePath");
        }

        // 22245: Support % and # in files and folders
        // StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativeUrl(@v)/Versions/DeleteAll?@v='" + Util.encodeUrl(filePath) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativePath(decodedUrl=@v)/Versions/DeleteAll?@v='" + Util.encodeUrl(Util.escapeQueryUrl(filePath)) + "'");
        if (callback.isDebug()) {
            callback.printDebug("deleteAllFileVersions", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("DeleteAll");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), null, "DELETE", handler);
        return handler.isSuccess();
    }

    /**
     * Delete file version.
     *
     * @param filePath     the file path
     * @param versionLabel the version label
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean deleteFileVersion(String siteUrl, String filePath, String versionLabel) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (filePath == null) {
            throw new IllegalArgumentException("filePath");
        }

        // 22245: Support % and # in files and folders
        // StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativeUrl(@v)/Versions/DeleteByLabel(versionlabel=" + versionLabel + ")?@v='" + Util.encodeUrl(filePath) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativePath(decodedUrl=@v)/Versions/DeleteByLabel(versionlabel=" + versionLabel + ")?@v='" + Util.encodeUrl(Util.escapeQueryUrl(filePath)) + "'");
        if (callback.isDebug()) {
            callback.printDebug("deleteFileVersion", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("DeleteByLabel");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), null, "DELETE", handler);
        return handler.isSuccess();
    }

    /**
     * Restore file version.
     *
     * @param filePath     the file path
     * @param versionLabel the version label
     * @return true, if successful
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean restoreFileVersion(String siteUrl, String filePath, String versionLabel) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (filePath == null) {
            throw new IllegalArgumentException("filePath");
        }

        // 22245: Support % and # in files and folders
        // StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativeUrl(@v)/Versions/RestoreByLabel(versionlabel=" + versionLabel + ")?@v='" + Util.encodeUrl(filePath) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativePath(decodedUrl=@v)/Versions/RestoreByLabel(versionlabel=" + versionLabel + ")?@v='" + Util.encodeUrl(Util.escapeQueryUrl(filePath)) + "'");
        if (callback.isDebug()) {
            callback.printDebug("restoreFileVersion", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.BooleanHandler handler = new ServiceResponseUtil.BooleanHandler("RestoreByLabel");
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.isSuccess();
    }

    /**
     * Gets the folder.
     *
     * @param folderPath the folder path
     * @return the folder
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public Folder getFolder(String siteUrl, String folderPath) throws ServiceException {
        return getFolder(siteUrl, folderPath, null);
    }

    public Folder getFolder(String siteUrl, String folderPath, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (folderPath == null) {
            throw new IllegalArgumentException("folderPath");
        }

        // 22245: Support % and # in files and folders
        // StringBuilder requestUrl = new StringBuilder("_api/web/GetFolderByServerRelativeUrl(@v)?@v='" + Util.encodeUrl(folderPath) + "'");
        // String queryOptionsString = queryOptions != null ? Util.queryOptionsToString(queryOptions, true) : "";
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFolderByServerRelativePath(decodedUrl=@v)?@v='" + Util.encodeUrl(Util.escapeQueryUrl(folderPath)) + "'");
        if (queryOptions != null) {
            StringBuilder sbQuery = new StringBuilder("");
            Util.queryOptionsToString(sbQuery, queryOptions, true);
            requestUrl.append(sbQuery);
        }
        if (callback.isDebug()) {
            callback.printDebug("getFolder", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.FolderHandler handler = new ServiceResponseUtil.FolderHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getFolder();
    }

    /**
     * Update folder.
     *
     * @param folder the folder
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public void updateFolder(String siteUrl, Folder folder) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (folder == null) {
            throw new IllegalArgumentException("folder");
        }
        if (folder.getServerRelativeUrl() == null) {
            throw new IllegalArgumentException("The ServerRelativeUrl property must be set in order to update folder.");
        }

        // 22245: Support % and # in files and folders
        // StringBuilder requestUrl = new StringBuilder("_api/web/GetFolderByServerRelativeUrl(@v)?@v='" + Util.encodeUrl(folder.getServerRelativeUrl()) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFolderByServerRelativePath(decodedUrl=@v)?@v='" + Util.encodeUrl(Util.escapeQueryUrl(folder.getServerRelativeUrl())) + "'");
        String requestBody = folder.toString();
        if (callback.isDebug()) {
            callback.printDebug("updateFolder", siteUrl, requestUrl.toString(), requestBody);
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, "MERGE", null);
    }

    /**
     * Recycle folder.
     *
     * @param folderPath the folder path
     * @return the string
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public String recycleFolder(String siteUrl, String folderPath) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (folderPath == null) {
            throw new IllegalArgumentException("folderPath");
        }

        // 22245: Support % and # in files and folders
        // StringBuilder requestUrl = new StringBuilder("_api/web/GetFolderByServerRelativeUrl(@v)/recycle?@v='" + Util.encodeUrl(folderPath) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFolderByServerRelativePath(decodedUrl=@v)/recycle?@v='" + Util.encodeUrl(Util.escapeQueryUrl(folderPath)) + "'");
        if (callback.isDebug()) {
            callback.printDebug("recycleFolder", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.RecycleHandler handler = new ServiceResponseUtil.RecycleHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.getRecycle();
    }

    /**
     * Creates the list.
     *
     * @param list the list
     * @return the com.independentsoft.share. list
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    /*
    public com.independentsoft.share.List createList(String siteUrl, com.independentsoft.share.List list) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (list == null) {
            throw new IllegalArgumentException("list");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists");
        String requestBody = list.toCreateJSon();
        if (callback.isDebug()) {
            callback.printDebug("createList", siteUrl, requestUrl.toString(), requestBody);
        }
        ServiceResponseUtil.ListHandler handler = new ServiceResponseUtil.ListHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, handler);
        return handler.getList();
    }
    */

    /**
     * Update list.
     *
     * @param listId the list id
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public void updateList(String siteUrl, String listId, String requestBody) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')");
        // String requestBody = list.toUpdateJSon();
        if (callback.isDebug()) {
            callback.printDebug("updateList", siteUrl, requestUrl.toString(), requestBody);
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, "MERGE", "*", null);
    }

    /**
     * Creates the list item.
     *
     * @param list     the list
     * @param listItem the list item
     * @return the list item
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public ListItem createListItem(String siteUrl,
                                   com.independentsoft.share.List list,
                                   ListItem listItem,
                                   Map<String, FieldValue> mFieldValue,
                                   List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (list == null) {
            throw new IllegalArgumentException("list");
        }
        if (listItem == null) {
            throw new IllegalArgumentException("listItem");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + list.getId() + "')/items");
        requestUrl.append(sbQuery);
        // com.independentsoft.share.List list = getList(siteUrl, listId);
        String requestBody = listItem.toCreateJSon(list.getListItemEntityTypeFullName(), mFieldValue);
        if (callback.isDebug()) {
            callback.printDebug("createListItem", siteUrl, requestUrl.toString(), requestBody);
        }
        ServiceResponseUtil.ListItemHandler handler = new ServiceResponseUtil.ListItemHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, handler);
        ListItem newItem = handler.getListItem();
        if (!listItem.isFolder()) {
            return newItem;
        }
        // these field only can set by update after create
        Map<String, FieldValue> mFolderFieldValue = new LinkedHashMap<String, FieldValue>();
        mFolderFieldValue.put("FileSystemObjectType", new FieldValue.StringValue(listItem.getFileSystemObjectType().getValue()));
        setListItemFieldValues(siteUrl, list, newItem.getId(), mFolderFieldValue);
        return getListItem(siteUrl, list.getId(), newItem.getId(), queryOptions);
    }

    /**
     * Sets the field values.
     *
     * @param list        the list
     * @param listItemId  the list item id
     * @param mFieldValue the field values
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public void setListItemFieldValues(String siteUrl, com.independentsoft.share.List list, int listItemId, Map<String, FieldValue> mFieldValue) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (list == null) {
            throw new IllegalArgumentException("list");
        }
        if (mFieldValue == null) {
            throw new IllegalArgumentException("fieldValues");
        }
        if (mFieldValue.size() == 0) {
            throw new IllegalArgumentException("Collection fieldValues is emtpy");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + list.getId() + "')/items(" + listItemId + ")");
        String requestBody = new FieldValue.ObjectValue(list.getListItemEntityTypeFullName(), mFieldValue).toString();
        if (callback.isDebug()) {
            callback.printDebug("setListItemFieldValues", siteUrl, requestUrl.toString(), requestBody);
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, "MERGE", "*", null);
    }

    // [Start] 23535: Support another way to update document field value
    public void setListDocumentFieldValuesNoVerChange(String siteUrl, com.independentsoft.share.List list, int listItemId, Map<String, String> mValue) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (list == null) {
            throw new IllegalArgumentException("list");
        }
        if (mValue == null) {
            throw new IllegalArgumentException("fieldValues");
        }
        if (mValue.size() == 0) {
            throw new IllegalArgumentException("Collection values is emtpy");
        }

        SerializableJSONArray ja = new SerializableJSONArray();
        for (Map.Entry<String, String> entry : mValue.entrySet()) {
            SerializableJSONObject o = new SerializableJSONObject();
            o.put("FieldName", entry.getKey());
            o.put("FieldValue", entry.getValue());
            ja.put(o.getObject());
        }
        SerializableJSONObject o = new SerializableJSONObject();
        o.put("formValues", ja.getObject());
        o.put("bNewDocumentUpdate", true);

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + list.getId() + "')/items(" + listItemId + ")/file/listitemallfields/ValidateUpdateListItem");
        String requestBody = o.toString();
        if (callback.isDebug()) {
            callback.printDebug("setListDocumentFieldValuesNoVerChange", siteUrl, requestUrl.toString(), requestBody);
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, null, null, null);
    }
    // [End] 23535

    public void addQuickLaunch(String siteUrl, NavigationNode navigationNode) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/Navigation/QuickLaunch");
        String requestBody = navigationNode.toCreateJSon();
        if (callback.isDebug()) {
            callback.printDebug("addQuickLaunch", siteUrl, requestUrl.toString(), requestBody);
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, null);
    }

    public void addTopNavigationBar(String siteUrl, NavigationNode navigationNode) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/Navigation/TopNavigationbar");
        String requestBody = navigationNode.toCreateJSon();
        if (callback.isDebug()) {
            callback.printDebug("addTopNavigationBar", siteUrl, requestUrl.toString(), requestBody);
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, null);
    }

    public void deleteQuickLaunch(String siteUrl, String nodeId) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/Navigation/QuickLaunch(" + nodeId + ")");
        if (callback.isDebug()) {
            callback.printDebug("deleteQuickLaunch", siteUrl, requestUrl.toString());
        }
        doSendRequest(siteUrl, "DELETE", requestUrl.toString(), null);
    }

    public void deleteTopNavigationBar(String siteUrl, int nodeId) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/Navigation/TopNavigationbar(" + nodeId + ")");
        if (callback.isDebug()) {
            callback.printDebug("deleteTopNavigationBar", siteUrl, requestUrl.toString());
        }
        doSendRequest(siteUrl, "DELETE", requestUrl.toString(), null);
    }

    public void addNavigationChild(String siteUrl, int nodeId, NavigationNode navigationNode) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (nodeId < 0) {
            throw new IllegalArgumentException("The parameter nodeId must be non-negative.");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/Navigation/GetNodeById(" + nodeId + ")/Children");
        if (callback.isDebug()) {
            callback.printDebug("addNavigationChild", siteUrl, requestUrl.toString());
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), navigationNode.toCreateJSon(), null);
    }

    public void changeMasterPage(String siteUrl, String masterPageURL, String customMasterPageURL) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web");
        StringBuilder requestBody = new StringBuilder("{ '__metadata': { 'type': 'SP.Web'}");
        if (masterPageURL != null && masterPageURL.length() > 0) {
            requestBody.append(", 'MasterUrl' : '" + Util.encodeJson(masterPageURL) + "'");
        }
        if (customMasterPageURL != null && customMasterPageURL.length() > 0) {
            requestBody.append(", 'CustomMasterUrl' : '" + Util.encodeJson(customMasterPageURL) + "'");
        }
        requestBody.append("}");
        if (callback.isDebug()) {
            callback.printDebug("changeMasterPage", siteUrl, requestUrl.toString(), requestBody.toString());
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody.toString(), "MERGE", null);
    }

    /*
    public ListItem getListItemByServerRelativeUrl(String siteUrl, String serverRelativeUrl) throws ServiceException {
        return getListItemByServerRelativeUrl(siteUrl, serverRelativeUrl, null);
    }

    public ListItem getListItemByServerRelativeUrl(String siteUrl, String serverRelativeUrl, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (serverRelativeUrl == null) {
            throw new IllegalArgumentException("serverRelativeUrl");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        queryOptionsString = queryOptionsString.isEmpty() ? queryOptionsString : "&" + queryOptionsString.substring(1);
        // 22245: Support % and # in files and folders
        // StringBuilder requestUrl = new StringBuilder("_api/web/GetFolderByServerRelativeUrl(@v)/listitemallfields?@v='" + Util.encodeUrl(serverRelativeUrl) + "'" + queryOptionsString;
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFolderByServerRelativePath(decodedUrl=@v)/listitemallfields?@v='" + Util.encodeUrl(serverRelativeUrl) + "'" + queryOptionsString;
        if (callback.isDebug()) {
            callback.printDebug("getListItemByServerRelativeUrl", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ListItemHandler handler = new ServiceResponseUtil.ListItemHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getListItem();
    }

    public List<ListItem> getListItemsByListTitle(String siteUrl, String listTitle) throws ServiceException {
        return getListItemsByListTitle(siteUrl, listTitle, null);
    }

    public List<ListItem> getListItemsByListTitle(String siteUrl, String listTitle, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listTitle == null) {
            throw new IllegalArgumentException("listTitle");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists/getbytitle('" + Util.encodeUrl(listTitle) + "')/items" + queryOptionsString;
        if (callback.isDebug()) {
            callback.printDebug("getListItemsByListTitle", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ListItemsHandler handler = new ServiceResponseUtil.ListItemsHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getListItems();
    }
    */

    // https://social.technet.microsoft.com/Forums/systemcenter/en-US/9bb56d0a-c199-488e-a491-8db245b1a22f/sharepoint-2013-office-365-rest-api-filter-by-filedirref?forum=sharepointgeneral
    // There is a bug in the filtering feature in sharepoint api. By applying caml query it can workaround.
    /*
    public List<ListItem> getListItemsCAML(String listId, List<IQueryOption> queryOptions, String sCamlQuery) throws ServiceException {
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }

        InputStream inputStream = null;
        List<ListItem> items = new ArrayList<ListItem>();
        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/getitems" + queryOptionsString;

        String requestBody = sCamlQuery;

        try {
            inputStream = sendRequest("POST", requestUrl.toString(), requestBody);

            //    debug(inputStream);

            items = parseGetListItems(inputStream);
        } catch (ServiceException e) {
            throw e;
        } catch (Exception e) {
            throw new ServiceException(e.getMessage(), e, requestUrl.toString(), requestBody);
        } finally {
            if (inputStream != null) {
                try {
                    inputStream.close();
                } catch (IOException e) {
                    throw new ServiceException(e.getMessage(), e, requestUrl.toString(), requestBody);
                }
            }

            if (httpClient != null && connectionManager == null) {
                try {
                    httpClient.close();
                } catch (IOException e) {
                    throw new ServiceException(e.getMessage(), e, requestUrl.toString());
                }
            }
        }

        return items;
    }
    */
    /*
    private ListItemCollection getListItemCollectionCAML(String siteUrl, String listId, List<IQueryOption> queryOptions, String sCamlQuery) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/getitems" + queryOptionsString;
        String requestBody = sCamlQuery;
        if (callback.isDebug()) {
            callback.printDebug("getListItemCollectionCAML", siteUrl, requestUrl.toString(), requestBody);
        }
        ServiceResponseUtil.ListItemCollectionHandler handler = new ServiceResponseUtil.ListItemCollectionHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, handler);
        return handler.getListItemCollection();
    }

    private ListItemCollection getListItemCollection(String siteUrl, String listId) throws ServiceException {
        return getListItemCollection(siteUrl, listId, null);
    }

    public ListItemCollection getListItemCollection(String siteUrl, String listId, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/items" + queryOptionsString;
        if (callback.isDebug()) {
            callback.printDebug("getListItemCollection", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ListItemCollectionHandler handler = new ServiceResponseUtil.ListItemCollectionHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getListItemCollection();
    }

    private ListItemCollection getListItemCollectionWithQuery(String siteUrl, String listId, String queryOptionsString) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }

        //  String queryOptionsString = Util.queryOptionsToString(queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/items" + queryOptionsString;
        if (callback.isDebug()) {
            callback.printDebug("getListItemCollectionWithQuery", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ListItemCollectionHandler handler = new ServiceResponseUtil.ListItemCollectionHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getListItemCollection();
    }

    private ListItemCollection getListItemCollectionCAMLWithQuery(String siteUrl, String listId, String queryOptionsString, String sCamlQuery) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }

        // String queryOptionsString = Util.queryOptionsToString(queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/getitems" + queryOptionsString;
        String requestBody = sCamlQuery;
        if (callback.isDebug()) {
            callback.printDebug("getListItemCollectionCAMLWithQuery", siteUrl, requestUrl.toString(), requestBody);
        }
        ServiceResponseUtil.ListItemCollectionHandler handler = new ServiceResponseUtil.ListItemCollectionHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, handler);
        return handler.getListItemCollection();
    }
    */

    public Folder createListFolder(String siteUrl, String serverRelativeUrl, String folderName) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (serverRelativeUrl == null) {
            throw new IllegalArgumentException("serverRelativeUrl");
        }
        if (folderName == null) {
            throw new IllegalArgumentException("folderName");
        }

        // 22245: Support % and # in files and folders
        // StringBuilder requestUrl = new StringBuilder("_api/web/GetFolderByServerRelativeUrl(@a)/AddSubFolder(@b)?@a='" + Util.encodeUrl(serverRelativeUrl) + "'" + "&@b='" + Util.encodeUrl(folderName) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFolderByServerRelativePath(decodedUrl=@a)/AddSubFolder(@b)?@a='" + Util.encodeUrl(Util.escapeQueryUrl(serverRelativeUrl)) + "'" + "&@b='" + Util.encodeUrl(Util.escapeQueryUrl(folderName)) + "'");
        if (callback.isDebug()) {
            callback.printDebug("createListFolder", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.FolderHandler handler = new ServiceResponseUtil.FolderHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.getFolder();
    }

    public Folder getFolderByName(String siteUrl, String parentPath, String folderName) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (parentPath == null) {
            throw new IllegalArgumentException("parentPath");
        }
        if (folderName == null) {
            throw new IllegalArgumentException("folderName");
        }

        // 22245: Support % and # in files and folders
        // StringBuilder requestUrl = new StringBuilder("_api/web/GetFolderByServerRelativeUrl(@v)/Folders(@v2)?@v='" + Util.encodeUrl(parentPath) + "'&@v2='" + Util.encodeUrl(folderName) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFolderByServerRelativePath(decodedUrl=@v)/Folders(@v2)?@v='" + Util.encodeUrl(Util.escapeQueryUrl(parentPath)) + "'&@v2='" + Util.encodeUrl(Util.escapeQueryUrl(folderName)) + "'");
        if (callback.isDebug()) {
            callback.printDebug("getFolderByName", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.FolderHandler handler = new ServiceResponseUtil.FolderHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getFolder();
    }

    public void activateSiteFeature(String siteUrl, String featureId, boolean bRetryable) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (featureId == null) {
            throw new IllegalArgumentException("featureId");
        }

        StringBuilder requestUrl = new StringBuilder("_api/site/features/add('" + featureId + "')");
        if (callback.isDebug()) {
            callback.printDebug("activateSiteFeature", siteUrl, requestUrl.toString());
        }

        // [Start] 23940: Changed to handle exception in this class
        try {
            doSendRequest(siteUrl, "POST", requestUrl.toString(), new ServiceResponseUtil.SiteFeatureHandler());
        } catch (ServiceException e) {
            if (e.getCause() instanceof SocketTimeoutException) {
                for (int iRetryCount = 0; iRetryCount < 30; iRetryCount++) {
                    try {
                        Thread.sleep(20 * 1000);
                        if (isSiteFeatureActivated(siteUrl, featureId)) {
                            return;
                        }
                    } catch (Throwable t) {
                        // continue
                    }
                }
            }
            throw e;
        }
        // [End] 23940
    }

    public void deactivateSiteFeature(String siteUrl, String featureId) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (featureId == null) {
            throw new IllegalArgumentException("featureId");
        }

        StringBuilder requestUrl = new StringBuilder("_api/site/features/remove('" + featureId + "')");
        if (callback.isDebug()) {
            callback.printDebug("deactivateSiteFeature", siteUrl, requestUrl.toString());
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), null);
    }

    public void activateWebFeature(String siteUrl, String featureId) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (featureId == null) {
            throw new IllegalArgumentException("featureId");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/features/add('" + featureId + "')");
        if (callback.isDebug()) {
            callback.printDebug("activateWebFeature", siteUrl, requestUrl.toString());
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), null);
    }

    public void deactivateWebFeature(String siteUrl, String featureId) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (featureId == null) {
            throw new IllegalArgumentException("featureId");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/features/remove('" + featureId + "')");
        if (callback.isDebug()) {
            callback.printDebug("deactivateWebFeature", siteUrl, requestUrl.toString());
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), null);
    }

    public void deleteListItemAttachment(String siteUrl, String listId, int itemId, String attachmentName) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (attachmentName == null) {
            throw new IllegalArgumentException("attachmentName");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/items(" + itemId + ")/AttachmentFiles/getByFileName('" + Util.encodeUrl(attachmentName) + "')");
        if (callback.isDebug()) {
            callback.printDebug("deleteListItemAttachment", siteUrl, requestUrl.toString());
        }
        doSendRequest(siteUrl, "DELETE", requestUrl.toString(), null);
    }

    public File addTemplatePage(String siteUrl, String folderRelativeUrl, String fileRelativeUrl, TemplateFileType templateFileType) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (folderRelativeUrl == null) {
            throw new IllegalArgumentException("folderRelativeUrl");
        }
        if (fileRelativeUrl == null) {
            throw new IllegalArgumentException("fileRelativeUrl");
        }
        if (templateFileType == null) {
            throw new IllegalArgumentException("templateFileType");
        }

        // 22245: Support % and # in files and folders
        // StringBuilder requestUrl = new StringBuilder("_api/web/GetFolderByServerRelativeUrl(@v)/Files/AddTemplateFile(urlOfFile=@v2,templateFileType=" + templateFileType.getValue() + ")?@v='" + Util.encodeUrl(folderRelativeUrl) + "'&@v2='" + Util.encodeUrl(fileRelativeUrl) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFolderByServerRelativePath(decodedUrl=@v)/Files/AddTemplateFile(urlOfFile=@v2,templateFileType=" + templateFileType.getValue() + ")?@v='" + Util.encodeUrl(Util.escapeQueryUrl(folderRelativeUrl)) + "'&@v2='" + Util.encodeUrl(Util.escapeQueryUrl(fileRelativeUrl)) + "'");
        if (callback.isDebug()) {
            callback.printDebug("addTemplatePage", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.FileHandler handler = new ServiceResponseUtil.FileHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.getFile();
    }

    public String getEtagOfListItem(String siteUrl, String fileRelativeUrl) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (fileRelativeUrl == null) {
            throw new IllegalArgumentException("fileRelativeUrl");
        }

        // 22245: Support % and # in files and folders
        // StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativeUrl(@v)/ListItemAllFields?$select=GUID&@v='" + Util.encodeUrl(fileRelativeUrl) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativePath(decodedUrl=@v)/ListItemAllFields?$select=GUID&@v='" + Util.encodeUrl(Util.escapeQueryUrl(fileRelativeUrl)) + "'");
        if (callback.isDebug()) {
            callback.printDebug("getEtagOfListItem", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ListItemEtagHandler handler = new ServiceResponseUtil.ListItemEtagHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.getListItemEtag();
    }

    public void updateWikiPage(String siteUrl, String listId, String fileRelativeUrl, String content) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (fileRelativeUrl == null) {
            throw new IllegalArgumentException("fileRelativeUrl");
        }
        if (content == null) {
            throw new IllegalArgumentException("content");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }

        // 22245: Support % and # in files and folders
        // StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativeUrl(@v)/ListItemAllFields?@v='" + Util.encodeUrl(fileRelativeUrl) + "'");
        StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativePath(decodedUrl=@v)/ListItemAllFields?@v='" + Util.encodeUrl(Util.escapeQueryUrl(fileRelativeUrl)) + "'");
        com.independentsoft.share.List list = getList(siteUrl, listId);
        String requestBody = "{\n" +
                " \"__metadata\": { \"type\": \"" + list.getListItemEntityTypeFullName() + "\" },\n" +
                " \"WikiField\" : \"" + Util.encodeJson(content) + "\"\n" +
                "}";
        String eTag = getEtagOfListItem(siteUrl, fileRelativeUrl);
        if (callback.isDebug()) {
            callback.printDebug("updateWikiPage", "siteUrl: " + siteUrl + ", requestUrl: " + requestUrl + ", requestBody: " + requestBody + ", eTag: " + eTag);
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, "MERGE", eTag, null);
    }

    // [Start] 23940: Added method to check whether a site feature is activated
    public boolean isSiteFeatureActivated(String siteUrl, String featureId) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (featureId == null) {
            throw new IllegalArgumentException("featureId");
        }

        StringBuilder requestUrl = new StringBuilder("_api/site/features/getById('" + featureId + "')?$select=DisplayName,DefinitionId");
        if (callback.isDebug()) {
            callback.printDebug("isSiteFeatureActivated", siteUrl, requestUrl.toString(), featureId);
        }
        ServiceResponseUtil.SiteFeatureHandler handler = new ServiceResponseUtil.SiteFeatureHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.isActivated();
    }
    // [End] 23940

    public List<Feature> getSiteFeatures(String siteUrl) throws ServiceException {
        return getSiteFeatures(siteUrl, null);
    }

    public List<Feature> getSiteFeatures(String siteUrl, List<IQueryOption> queryOptions) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder sbQuery = new StringBuilder("");
        Util.queryOptionsToString(sbQuery, queryOptions);
        StringBuilder requestUrl = new StringBuilder("_api/site/Features");
        requestUrl.append(sbQuery);
        if (callback.isDebug()) {
            callback.printDebug("getSiteFeatures", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.FeaturesHandler handler = new ServiceResponseUtil.FeaturesHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getFeatures();
    }

    public List<Locale> getSiteSupportedUILanguage(String siteUrl) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/SupportedUILanguageIds");
        if (callback.isDebug()) {
            callback.printDebug("getSiteSupportedUILanguage", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.SupportedLocalesHandler handler = new ServiceResponseUtil.SupportedLocalesHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getSupportedLocales();
    }

    /**
     * Creates the list items by batch.
     *
     * @param listId    the list id
     * @param listItems the list items
     * @return the list items that successfully created
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public List<ListItem> createListItems(String siteUrl, String listId, List<String> listItems) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/$batch");
        String createListItemRequestUrl = "_api/web/lists('" + listId + "')/items";
        String serviceUrl = siteUrl.endsWith("/") ? siteUrl + createListItemRequestUrl : siteUrl + "/" + createListItemRequestUrl;
        String sBatchBoundary = UUID.randomUUID().toString();
        String sChangeSetBoundary = UUID.randomUUID().toString();
        String sCommonChangeSetBoundary = "--changeset_" + sChangeSetBoundary + "\r\n" +
                "Content-Type: application/http" + "\r\n" +
                "Content-Transfer-Encoding: binary" + "\r\n\r\n" +
                "POST " + serviceUrl + " HTTP/1.1" + "\r\n" +
                "Content-Type: application/json;odata=verbose";
        String sBatchBody = "";
        for (String slistItem : listItems) {
            sBatchBody += sCommonChangeSetBoundary + "\r\n\r\n" +
                    slistItem + "\r\n\r\n";
        }
        String requestBody = "--batch_" + sBatchBoundary + "\r\n" +
                "Content-Type: multipart/mixed; boundary=\"changeset_" + sChangeSetBoundary + "\"" + "\r\n" +
                "Content-Transfer-Encoding: binary" + "\r\n\r\n" +
                sBatchBody +
                "--changeset_" + sChangeSetBoundary + "--" + "\r\n\r\n" +
                "--batch_" + sBatchBoundary + "--";
        if (callback.isDebug()) {
            callback.printDebug("createListItems", siteUrl, requestUrl.toString(), requestBody);
        }
        ServiceResponseUtil.CreatedListItemsHandler handler = new ServiceResponseUtil.CreatedListItemsHandler();
        doSendBatchBoundaryRequest(siteUrl, "POST", requestUrl.toString(), requestBody, sBatchBoundary, handler);
        return handler.getListItems();
    }

    /**
     * Delete list items by batch.
     *
     * @param listId    the list id
     * @param listItems the list items
     * @return is all items delete successfully
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public boolean deleteListItems(String siteUrl, String listId, List<ListItem> listItems) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/$batch");
        String createListItemRequestUrl = "_api/web/lists('" + listId + "')/items";
        String serviceUrl = siteUrl.endsWith("/") ? siteUrl + createListItemRequestUrl : siteUrl + "/" + createListItemRequestUrl;
        String sBatchBoundary = UUID.randomUUID().toString();
        String sChangeSetBoundary = UUID.randomUUID().toString();
        String sCommonChangeSetBoundary = "--changeset_" + sChangeSetBoundary + "\r\n" +
                "Content-Type: application/http" + "\r\n" +
                "Content-Transfer-Encoding: binary" + "\r\n\r\n";
        String sBatchBody = "";
        for (ListItem listItem : listItems) {
            sBatchBody += sCommonChangeSetBoundary +
                    "DELETE " + serviceUrl + "(" + listItem.getId() + ")" + " HTTP/1.1" + "\r\n" +
                    "If-Match: *" + "\r\n\r\n";
        }
        String requestBody = "--batch_" + sBatchBoundary + "\r\n" +
                "Content-Type: multipart/mixed; boundary=\"changeset_" + sChangeSetBoundary + "\"" + "\r\n" +
                "Content-Transfer-Encoding: binary" + "\r\n\r\n" +
                sBatchBody +
                "--changeset_" + sChangeSetBoundary + "--" + "\r\n\r\n" +
                "--batch_" + sBatchBoundary + "--";
        if (callback.isDebug()) {
            callback.printDebug("deleteListItems", siteUrl, requestUrl.toString(), requestBody);
        }
        ServiceResponseUtil.DeletedListItemsHandler handler = new ServiceResponseUtil.DeletedListItemsHandler();
        doSendBatchBoundaryRequest(siteUrl, "POST", requestUrl.toString(), requestBody, sBatchBoundary, handler);
        return handler.isSuccess();
    }

    /**
     * Gets the list content types.
     *
     * @param listId the list id
     * @throws com.independentsoft.share.ServiceException the service exception
     */
    public String addListContentTypes(String siteUrl, String listId, String contentTypeId) throws ServiceException {
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (contentTypeId == null) {
            throw new IllegalArgumentException("contentTypeId");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/contenttypes/addAvailableContentType");
        String requestBody = "{'contentTypeId':'" + contentTypeId + "'}";
        if (callback.isDebug()) {
            callback.printDebug("addListContentTypes", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.ContentTypeHandler handler = new ServiceResponseUtil.ContentTypeHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, handler);
        return handler.getContentType().getId();
    }

    private void loadAvailableListSettingsLinks(String siteUrl, String listId, boolean hasUniqueRoleAssignments, ListSettings settings)
            throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (settings == null) {
            throw new IllegalArgumentException("settings");
        }

        String urlId = "?List=" + Util.encodeUrl("{" + listId + "}");
        ServiceResponseUtil.StringHandler handler = new ServiceResponseUtil.StringHandler();
        doSendRequest(siteUrl, "GET", "/_layouts/15/listedit.aspx" + urlId, handler);
        try {
            settings.setAvailableLinks(handler.getString(), hasUniqueRoleAssignments);
        } catch (Throwable t) {
            if (t instanceof ServiceException) {
                throw (ServiceException) t;
            }
            throw new ServiceException(t.getMessage(), t);
        }
    }

    private String getListSettingBody(String siteUrl, String listId, ListSettings.GetLink link)
            throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (link == null) {
            throw new IllegalArgumentException("link");
        }

        String urlId = "?List=" + Util.encodeUrl("{" + listId + "}");
        ServiceResponseUtil.StringHandler handler = new ServiceResponseUtil.StringHandler();
        try {
            if (link instanceof ListSettings.GetLink.Html) {
                String sHtml = "";
                String url = ((ListSettings.GetLink.Html) link).getUrl();
                if (url != null && !"".equals(url)) {
                    url += urlId;
                    if (callback.isDebug()) {
                        callback.printDebug("getListHtmlSettings", siteUrl, url);
                    }
                    try {
                        doSendRequest(siteUrl, "GET", url, handler);
                        sHtml = handler.getString();
                    } catch (Throwable t) {
                        // Access Denied will be in here
                    }
                }
                return sHtml;
            }
            if (link instanceof ListSettings.GetLink.Xml) {
                // String queryOptionsString = Util.queryOptionsToString(((ListSettings.GetLink.Xml) link).getQueryOption());
                StringBuilder sbQuery = new StringBuilder("");
                Util.queryOptionsToString(sbQuery, ((ListSettings.GetLink.Xml) link).getQueryOption());
                StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')");
                requestUrl.append(sbQuery);
                if (callback.isDebug()) {
                    callback.printDebug("getListXmlSettings", siteUrl, requestUrl.toString());
                }
                doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
                return handler.getString();
            }
            if (link instanceof ListSettings.GetLink.FormInfoPath) {
                List<ContentType> contentTypes = getListContentTypes(siteUrl, listId);
                if (contentTypes == null) {
                    return null;
                }
                ArrayList<InfoPath> infoPaths = new ArrayList<InfoPath>();
                for (ContentType contentType : contentTypes) {
                    infoPaths.add(FormsServiceUtils.getInfoPath(this, siteUrl, listId, contentType));
                }
                InfoPathCollection infoPathCollection = new InfoPathCollection();
                infoPathCollection.addInfoPaths(infoPaths);
                return infoPathCollection.getDescriptor();
            }
            if (link instanceof ListSettings.GetLink.WorkflowSubscription) {
                return ProcessQueryUtils.getWorkflowSubscriptionsJson(this, siteUrl, listId);
            }

            return null;
        } catch (Throwable t) {
            if (t instanceof ServiceException) {
                throw (ServiceException) t;
            }
            throw new ServiceException(t.getMessage(), t);
        }
    }

    private void loadListSettings(String siteUrl, String listId, Map<String, ListSettings.GetLink> mLinks, boolean bRefresh, ListSettings settings)
            throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (mLinks == null) {
            throw new IllegalArgumentException("mLinks");
        }

        try {
            for (Map.Entry<String, ListSettings.GetLink> entry : mLinks.entrySet()) {
                String sType = entry.getKey();
                if (!bRefresh && settings.isPropertyExists(sType)) {
                    continue;
                }
                ListSettings.GetLink link = entry.getValue();
                String sResult = getListSettingBody(siteUrl, listId, link);
                if (sResult != null) {
                    settings.setAttribute(new RawEntity.AttributeType(sType), sResult);
                }
            }
        } catch (Throwable t) {
            if (t instanceof ServiceException) {
                throw (ServiceException) t;
            }
            throw new ServiceException(t.getMessage(), t);
        }
    }

    public Map<String, ListSettings.GetLink> getListSettingsAvailableLinks(String siteUrl, String listId, boolean hasUniqueRoleAssignments, ListSettings.Type type, ListSettings settings)
            throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (settings == null) {
            throw new IllegalArgumentException("settings");
        }

        try {
            Map<ListSettings.Type, LinkedHashMap<String, ListSettings.GetLink>> mAvailableLinks = settings.getAvailableLinks();
            if (mAvailableLinks == null) {
                loadAvailableListSettingsLinks(siteUrl, listId, hasUniqueRoleAssignments, settings);
                mAvailableLinks = settings.getAvailableLinks();
            }
            if (mAvailableLinks == null) {
                return null;
            }
            return mAvailableLinks.get(type);
        } catch (Throwable t) {
            throw new ServiceException(t.getMessage(), t);
        }
    }

    public ListSettings.Config getListSettingsConfig(String siteUrl, String listId, boolean hasUniqueRoleAssignments, ListSettings.Type type, boolean bRefresh, ListSettings settings)
            throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (type == null) {
            throw new IllegalArgumentException("type");
        }
        if (settings == null) {
            throw new IllegalArgumentException("settings");
        }

        try {
            Map<String, ListSettings.GetLink> mLinks = getListSettingsAvailableLinks(siteUrl, listId, hasUniqueRoleAssignments, type, settings);
            if (mLinks == null) {
                return null;
            }
            loadListSettings(siteUrl, listId, mLinks, bRefresh, settings);
            return settings.getConfig(type);
        } catch (Throwable t) {
            throw new ServiceException(t.getMessage(), t);
        }
    }

    public boolean breakListRoleInheritance(String siteUrl, String listId, boolean copyRoleAssignments) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/BreakRoleInheritance(" + Boolean.toString(copyRoleAssignments).toLowerCase() + ")");
        if (callback.isDebug()) {
            callback.printDebug("breakListRoleInheritanceImplementation", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.DeletedListItemsHandler handler = new ServiceResponseUtil.DeletedListItemsHandler();
        doSendRequest(siteUrl, "POST", requestUrl.toString(), handler);
        return handler.isSuccess();
    }

    private void updateListSettings(String siteUrl, String requestUrl, String requestBody)
            throws ServiceException {
        if (callback.isDebug()) {
            callback.printDebug("updateListSettings", siteUrl, requestUrl, requestBody);
        }
        //ServiceResponseUtil.StringHandler handler = new ServiceResponseUtil.StringHandler();
        doSendRefererRequest(
                siteUrl,
                "POST",
                requestUrl,
                requestBody,
                siteUrl + requestUrl,
                null
        );
        //System.out.println(handler.getString());
    }

    public boolean updateListSettings(String siteUrl,
                                      String listId,
                                      Map<String, ListSettings.GetLink> mGetLink,
                                      Map<String, ListSettings.PostLink> mPostLink,
                                      Map<String, ListSettings.PostLink> mCurrPostLink,
                                      Util.UpdateReporter reporter)
            throws ServiceException {
        boolean bUpdated = false;
        String urlId = "?List=" + Util.encodeUrl("{" + listId + "}");
        for (Map.Entry<String, ListSettings.PostLink> entry : mPostLink.entrySet()) {
            ListSettings.PostLink link = entry.getValue();
            try {
                if (link instanceof ListSettings.PostLink.Html) {
                    ListSettings.PostLink.Html html = (ListSettings.PostLink.Html) link;
                    ListSettings.GetLink srcLink = html.getSourceLink(mGetLink);
                    if (srcLink == null) {
                        continue;
                    }
                    String sCurrHtml = getListSettingBody(siteUrl, listId, srcLink);
                    ListSettings.PostLink.Html.RequestBody requestBody = html.getRequestBody(sCurrHtml);
                    if (requestBody == null) {
                        continue;
                    }
                    ArrayList<String> alMissingElements = requestBody.getMissingElements();
                    for (String sElement : alMissingElements) {
                        callback.logHide(sElement);
                    }
                    String sBody = requestBody.getBody();
                    if ("".equals(sBody)) {
                        continue;
                    }
                    ListSettings.PostLink currLink = mCurrPostLink.get(entry.getKey());
                    if (!(currLink instanceof ListSettings.PostLink.Html)) {
                        continue;
                    }
                    ListSettings.PostLink.Html.RequestBody currRequestBody = ((ListSettings.PostLink.Html) currLink).getRequestBody(sCurrHtml);
                    if (currRequestBody == null || sBody.equals(currRequestBody.getBody())) {
                        continue;
                    }
                    if (reporter != null && !bUpdated) {
                        reporter.logStart(null);
                    }
                    updateListSettings(siteUrl, html.getUrl(sCurrHtml) + urlId, sBody);
                    bUpdated = true;
                } else if (link instanceof ListSettings.PostLink.Xml) {
                    ListSettings.PostLink.Xml xml = (ListSettings.PostLink.Xml) link;
                    if (mCurrPostLink != null) {
                        ListSettings.PostLink currLink = mCurrPostLink.get(entry.getKey());
                        if (!(currLink instanceof ListSettings.PostLink.Xml)) {
                            continue;
                        }
                        if (xml.getBody().equals(((ListSettings.PostLink.Xml) currLink).getBody())) {
                            continue;
                        }
                    }
                    if (reporter != null && !bUpdated) {
                        reporter.logStart(null);
                    }
                    updateList(siteUrl, listId, xml.getBody());
                    bUpdated = true;
                } else if (link instanceof ListSettings.PostLink.RootFolderProperty) {
                    ListSettings.PostLink.RootFolderProperty property = (ListSettings.PostLink.RootFolderProperty) link;
                    if (mCurrPostLink != null) {
                        ListSettings.PostLink currProperty = mCurrPostLink.get(entry.getKey());
                        if (!(currProperty instanceof ListSettings.PostLink.RootFolderProperty)) {
                            continue;
                        }
                        if (property.getTag().equals(((ListSettings.PostLink.RootFolderProperty) currProperty).getTag())
                                && property.getValue().equals(((ListSettings.PostLink.RootFolderProperty) currProperty).getValue())) {
                            continue;
                        }
                    }
                    if (reporter != null && !bUpdated) {
                        reporter.logStart(null);
                    }
                    ProcessQueryUtils.setListRootFolderProperty(
                            this, siteUrl, listId, property.getTag(), property.getValue()
                    );
                    bUpdated = true;
                }
            } catch (Throwable t) {
                if (t instanceof ServiceException) {
                    throw (ServiceException) t;
                }
                throw new ServiceException(t.getMessage(), t);
            }
        }
        if (reporter != null && bUpdated) {
            reporter.logEnd();
        }
        return bUpdated;
    }

    public void moveFolderByPath(String siteUrl, String srcPath, String destPath) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (srcPath == null) {
            throw new IllegalArgumentException("srcPath");
        }
        if (destPath == null) {
            throw new IllegalArgumentException("destPath");
        }

        // POST https://ahsay.sharepoint.com/sites/R08_Restore/_api/SP.MoveCopyUtil.MoveFolderByPath() HTTP/1.1
        // {"srcPath":{"__metadata":{"type":"SP.ResourcePath"},"DecodedUrl":"https://ahsay.sharepoint.com/sites/R08_Restore/Lists/Test/Folder2"},"destPath":{"__metadata":{"type":"SP.ResourcePath"},"DecodedUrl":"https://ahsay.sharepoint.com/sites/R08_Restore/Lists/Test/Folder2"}}
        StringBuilder requestUrl = new StringBuilder("_api/SP.MoveCopyUtil.MoveFolderByPath()");
        String requestBody = "{\"" +
                "srcPath\":{\"__metadata\":{\"type\":\"SP.ResourcePath\"},\"DecodedUrl\":\"" + srcPath + "\"}" +
                ",\"destPath\":{\"__metadata\":{\"type\":\"SP.ResourcePath\"},\"DecodedUrl\":\"" + destPath + "\"}" +
                "}";
        if (callback.isDebug()) {
            callback.printDebug("moveFolderByPath", siteUrl, requestUrl.toString(), requestBody);
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, null);
    }

    public void moveFileByPath(String siteUrl, String srcPath, String destPath) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (srcPath == null) {
            throw new IllegalArgumentException("srcPath");
        }
        if (destPath == null) {
            throw new IllegalArgumentException("destPath");
        }

        // POST https://ahsay.sharepoint.com/sites/R08_Restore/_api/SP.MoveCopyUtil.MoveFileByPath() HTTP/1.1
        // {"srcPath":{"__metadata":{"type":"SP.ResourcePath"},"DecodedUrl":"https://ahsay.sharepoint.com/sites/R08_Restore/Lists/Test/117_.000"},"destPath":{"__metadata":{"type":"SP.ResourcePath"},"DecodedUrl":"https://ahsay.sharepoint.com/sites/R08_Restore/Lists/Test/117_.000"}}
        StringBuilder requestUrl = new StringBuilder("_api/SP.MoveCopyUtil.MoveFileByPath()");
        String requestBody = "{\"" +
                "srcPath\":{\"__metadata\":{\"type\":\"SP.ResourcePath\"},\"DecodedUrl\":\"" + srcPath + "\"}" +
                ",\"destPath\":{\"__metadata\":{\"type\":\"SP.ResourcePath\"},\"DecodedUrl\":\"" + destPath + "\"}" +
                "}";
        if (callback.isDebug()) {
            callback.printDebug("moveFileByPath", siteUrl, requestUrl.toString(), requestBody);
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), requestBody, null);
    }

    private ArrayList<WorkflowDefinition> getWorkflow2010Definitions(String sSiteUrl, String sFolder, boolean bDefault) throws ServiceException {
        String method = "list+documents";
        String command = "listHiddenDocs=true" +
                "&listExplorerDocs=false" +
                "&listRecurse=true" +
                "&listFiles=true" +
                "&listFolders=true" +
                "&listLinkInfo=true" +
                "&listIncludeParent=true" +
                "&listDerived=false" +
                "&listBorders=false" +
                "&listChildWebs=true" +
                "&listThickets=false" +
                "&initialUrl=" + Util.encodeUrl(sFolder);
        ServiceResponseUtil.StringHandler handler = new ServiceResponseUtil.StringHandler();
        doSendAuthorDllRequest(sSiteUrl, method, command, handler);

        ArrayList<WorkflowDefinition> alDefinition = new ArrayList<WorkflowDefinition>();
        try {
            ArrayList<String> alInput = Util.StringParser.getBoundaries("document_name", "<ul>", "</ul>", handler.getString());
            for (String input : alInput) {
                if (!input.contains("WorkflowDisplayName")) {
                    continue;
                }
                WorkflowDefinition.Custom2010 definition = bDefault ? new WorkflowDefinition.Default2010() : new WorkflowDefinition.Custom2010();
                definition.setAttribute(new RawEntity.AttributeType(WorkflowDefinition.ATTR_PROPERTIES), input);
                String sType = definition.getType();
                if (!sType.equals(WorkflowDefinition.DEFAULT_REUSE_TYPE) && !sType.equals(WorkflowDefinition.CUSTOM_REUSE_TYPE)) {
                    String xmlRelativePath = Util.Url.getRelativeUrl(sSiteUrl + definition.getXmlRelativePath());
                    InputStream is = getFileStream(sSiteUrl, xmlRelativePath);
                    try {
                        definition.setAttribute(new RawEntity.AttributeType(WorkflowDefinition.ATTR_ASSOCIATION),
                                WorkflowDefinition.Custom2010.parseAssociation(Util.convertInputStreamToString(is))
                        );
                    } finally {
                        is.close();
                    }
                }
                alDefinition.add(definition);
            }
        } catch (Throwable t) {
        }
        return alDefinition;
    }

    public ArrayList<WorkflowDefinition> getWorkflowDefinitions(String sSiteUrl) throws ServiceException {
        // ArrayList<WorkflowDefinition> alDefinition = getWorkflow2010Definitions(sSiteUrl, "_catalogs/wfpub", true);
        // alDefinition.addAll(getWorkflow2010Definitions(sSiteUrl, "Workflows", false));
        ArrayList<WorkflowDefinition> alDefinition = getWorkflow2010Definitions(sSiteUrl, "Workflows", false);
        try {
            alDefinition.addAll(ProcessQueryUtils.getWorkflow2013Definitions(this, sSiteUrl));
        } catch (Exception e) {
            throw new ServiceException(e.getMessage(), e);
        }
        Collections.sort(alDefinition, WorkflowDefinition.NameComparator);
        Collections.sort(alDefinition, WorkflowDefinition.TypeComparator);
        return alDefinition;
    }

    public void createFolderByAuthorDll(String sSiteUrl, String sPath) throws ServiceException {
        String method = "create+url-directories";
        String command = "urldirs=[[url=" + Util.encodeUrl(sPath) + "]]";
        doSendAuthorDllRequest(sSiteUrl, method, command, null);
    }

    public void createFileByAuthorDll(String sSiteUrl, String FilePath, InputStream is, boolean overwrite) throws ServiceException {
        String method = "put+document";
        String command = "keep_checked_out=false" +
                "&comment=";
        try {
            if (overwrite) {
                command += "&put_option=overwrite";
            }
            command += "&document=[document_name=" + FilePath + "]]\n" + Util.convertInputStreamToString(is);
        } catch (Throwable t) {
            if (t instanceof ServiceException) {
                throw (ServiceException) t;
            }
            throw new ServiceException(t.getMessage(), t);
        }
        doSendAuthorDllRequest(sSiteUrl, method, command, null);
    }

    public Role getPrincipalRoleDefinition(String siteUrl, int sPrincipalId) throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/roleassignments/GetByPrincipalId('" + sPrincipalId + "')/RoleDefinitionBindings");

        if (callback.isDebug()) {
            System.out.println("[Service.getPrincipalRoleDefinition] siteUrl: " + siteUrl + ", requestUrl: " + requestUrl);
        }

        ServiceResponseUtil.RoleHandler handler = new ServiceResponseUtil.RoleHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getRole();
    }

    public CurrentThemeInfo getCurrentThemeInfo(String siteUrl)
            throws ServiceException {
        CurrentThemeInfo currentThemeInfo = new CurrentThemeInfo();
        loadCurrentThemeInfo(siteUrl, currentThemeInfo);
        return currentThemeInfo;
    }

    public void loadCurrentThemeInfo(String siteUrl, CurrentThemeInfo currentThemeInfo)
            throws ServiceException {
        ServiceResponseUtil.StringHandler handler = new ServiceResponseUtil.StringHandler();
        if (callback.isDebug()) {
            callback.printDebug("getCurrentThemeInfo", siteUrl);
        }
        doSendRequest(siteUrl, "GET", "", handler);
        ThemeInfo themeInfo = getThemeInfo(siteUrl);
        Site site = getSite(siteUrl);
        try {
            currentThemeInfo.setAttribute(new RawEntity.AttributeType(CurrentThemeInfo.ATTR_SURL), siteUrl);
            currentThemeInfo.setAttribute(new RawEntity.AttributeType(CurrentThemeInfo.ATTR_HTML), handler.getString());
            currentThemeInfo.setAttribute(new RawEntity.AttributeType(CurrentThemeInfo.ATTR_BGURL), themeInfo.getThemeBackgroundImageUri());
            currentThemeInfo.setAttribute(new RawEntity.AttributeType(CurrentThemeInfo.ATTR_MURL), site.getMasterUrl());
            currentThemeInfo.setAttribute(new RawEntity.AttributeType(CurrentThemeInfo.ATTR_CMURL), site.getCustomMasterUrl());
        } catch (Throwable t) {
            if (t instanceof ServiceException) {
                throw (ServiceException) t;
            }
            throw new ServiceException(t.getMessage(), t);
        }
    }

    public void createList(String sSiteUrl, ListTemplateType type, String sTitle)
            throws Exception {
        String sRequest = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<ows:Batch OnError=\"Return\" Version=\"15.00.0.000\">" +
                "<Method ID=\"0,NewList\">" +
                "<SetVar Name=\"Cmd\">NewList</SetVar>" +
                "<SetVar Name=\"ListTemplate\">" + type.getValue() + "</SetVar>" +
                "<SetVar Name=\"Title\">" + sTitle + "</SetVar>" +
                "<SetVar Name=\"RootFolder\" />" +
                "<SetVar Name=\"LangID\">1033</SetVar>" +
                "</Method>" +
                "</ows:Batch>";
        ServiceInstance.RequestType requestType = new ServiceInstance.RequestType() {
            @Override
            protected Map<String, String> getHeaders() {
                Map<String, String> mHdr = new LinkedHashMap<String, String>();
                mHdr.put("Accept", "application/atom+xml");
                mHdr.put("Content-Type", "application/xml");
                return mHdr;
            }
        };
        doSendCustomizeRequest(
                sSiteUrl,
                "POST",
                "_vti_bin/owssvr.dll?Cmd=DisplayPost",
                sRequest,
                requestType,
                null
        );
    }

    // [Start] 22363: Added method to check site exist by calling getSite without retry
    public boolean isSiteExist(String sSiteUrl)
            throws Throwable {
        try {
            // 22014: Set null
            Site site = getSite(sSiteUrl, false, null);
            return site != null;
        } catch (Throwable t) {
            // Search message with "404" as site not exist, otherwise throw exception
            if (t.getMessage().contains("404")) {
                return false;
            }
            throw t;
        }
    }
    // [End] 22363

    // [Start] 22215: Support list item version control
    public String getListItemVersion(String siteUrl, String listId, int itemId, String versionId)
            throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (itemId < 0) {
            throw new IllegalArgumentException("The parameter itemId must be non-negative.");
        }
        if (versionId == null) {
            throw new IllegalArgumentException("versionId");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/items(" + itemId + ")/versions(" + versionId + ")");
        if (callback.isDebug()) {
            callback.printDebug("getListItemVersion", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.StringHandler handler = new ServiceResponseUtil.StringHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getString();
    }

    public void deleteListItemVersion(String siteUrl, String listId, int itemId, String versionId)
            throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (listId == null) {
            throw new IllegalArgumentException("listId");
        }
        if (itemId < 0) {
            throw new IllegalArgumentException("The parameter itemId must be non-negative.");
        }
        if (versionId == null) {
            throw new IllegalArgumentException("versionId");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/lists('" + listId + "')/items(" + itemId + ")/versions(" + versionId + ")");
        if (callback.isDebug()) {
            callback.printDebug("deleteListItemVersion", siteUrl, requestUrl.toString());
        }
        doSendRequest(siteUrl, "POST", requestUrl.toString(), null, "DELETE", "*", null);
    }
    // [End] 22215

    // [Start] 22869: Added method to get list item checkout user

    /**
     * @param siteUrl  the site url
     * @param filePath relative server file path
     * @return the user checked out the file, if exist
     * @throws Exception
     */
    public User getItemCheckedOutUser(String siteUrl, String filePath)
            throws ServiceException {
        if (siteUrl == null) {
            throw new IllegalArgumentException("siteUrl");
        }
        if (filePath == null) {
            throw new IllegalArgumentException("filePath");
        }

        StringBuilder requestUrl = new StringBuilder("_api/web/GetFileByServerRelativePath(decodedUrl=@v)/CheckedOutByUser");
        requestUrl.append("?@v='" + Util.encodeUrl(Util.escapeQueryUrl(filePath)) + "'");
        if (callback.isDebug()) {
            callback.printDebug("getItemCheckedOutUser", siteUrl, requestUrl.toString());
        }
        ServiceResponseUtil.UserHandler handler = new ServiceResponseUtil.UserHandler();
        doSendRequest(siteUrl, "GET", requestUrl.toString(), handler);
        return handler.getUser();
    }
    // [End] 22869
}