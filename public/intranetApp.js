'use strict';
var INTRANET = INTRANET || {};

ExecuteOrDelayUntilScriptLoaded(INTRANET.InIt, "sp.js");

var context;
var hostWebUrl, hostWebContext, hostWeb, hostWebContextFile, appWebUrl;
var countContentType = 0;
var countTotalMasterLookUpLists = 0;
var hostWebContentTypes;
var ctColumns = [];
var counterInstallation = 0;
var collAppStateItems;
var collContentTypes;
var collFields;
var collRemoveContentType = [];
var collRemoveFields = [];
var counterDeletedLists = 0;
var addFieldsPromise = $.Deferred();

$(window).load(function() {
    // executes when complete page is fully loaded, including all frames, objects and images

    var tempURL = window.location.href;
    tempURL = tempURL.replace("BravoConfig.aspx", "Home.aspx");
    $("#linkSiteURL").attr("href", tempURL);
    $("#idLinkBravoSite").attr("href", tempURL);

    var tempHostURL = decodeURIComponent(INTRANET.GetQueryStringParameter("SPHostUrl"));
    tempURL = tempHostURL + "/Cards/Forms/AllItems.aspx";
    $("#idCardsLibLink").attr("href", tempURL);

    tempURL = tempHostURL + "/Categories/Forms/AllItems.aspx";
    $("#idCategoriesLibLink").attr("href", tempURL);

    //tempURL = tempHostURL + "/BravoLensFilters/Forms/AllItems.aspx";
    //$("#idBravoLensFiltersLink").attr("href", tempURL);

    tempURL = tempHostURL + "/Lists/Department/AllItems.aspx";
    $("#idDepartmentLink").attr("href", tempURL);

    //var errDetails = decodeURIComponent(INTRANET.GetQueryStringParameter("ErrMsg"));
    //$("#errorDetails").innerText = errDetails;

    INTRANET.InIt();
    //INTRANET.GetAppStateItems();
});

INTRANET.App = {
    TotalSiteColumns: INTRANET.Schema.SiteColumns.length,
    TotalSiteContentTypes: INTRANET.Schema.ContentTypes.length,
    IsLookUpExecution: false,
    CollLists: [],
    ValidateLicenseWebAPI: "/api/ValidateLicense",
    CreateNewTenantAPI: "/api/CreateTenant",
    UploadSampleDataAPI: "/api/OnDemandCards/SampleCards",
    GetAllCardsAPI: "/api/OnDemandCards",
    AppWebURL: "",
    ListBravoAppState: "BravoAppState",
    Done: 'Done',
    APPState: { Initialization: 'Initialization', SiteConfiguration: 'Configuration', AppPermissions: 'AppPermissions', InstallationCompleted: 'InstallationCompleted' },
    SPGroup: { Name: 'BravoAdmin', Description: "'Use this group to grant people Admin permissions to the INTRANET portal where Admin can change configuration." },
    IsRollback: false,
    LicenseVerificationService: "https://verificationservice.officeapps.live.com/ova/verificationagent.svc/rest/verify?token=",
};

INTRANET.InIt = function() {

    INTRANET.Log(INTRANET.LogType.Info, "Execution Started....");

    context = SP.ClientContext.get_current();
    //$(document).ready(function () {
    hostWebUrl = decodeURIComponent(INTRANET.GetQueryStringParameter("SPHostUrl"));

    //Get Host Web Context using App web
    hostWebContext = new SP.AppContextSite(context, hostWebUrl);
    hostWeb = hostWebContext.get_web();
    //});
    INTRANET.Constant.HostWebURL = hostWebUrl;

    INTRANET.Constant.OrganizationName = $("#txtOrgName").val();
    INTRANET.Constant.ShortAddress = $("#txtShortAddress").val();
    INTRANET.Constant.AppWebURL = _spPageContextInfo.webAbsoluteUrl;
    //INTRANET.Constant.EmailDomain = INTRANET.RetrieveEmailDomain(_spPageContextInfo.userEmail);
    INTRANET.Constant.TenantId = _spPageContextInfo.aadTenantId;
};

INTRANET.CreateGuid = function() {
    function CreateNewGuid(s) {
        var p = (Math.random().toString(16) + "000000000").substr(2, 8);
        return s ? "-" + p.substr(0, 4) + "-" + p.substr(4, 4) : p;
    }
    return CreateNewGuid() + CreateNewGuid(true) + CreateNewGuid(true) + CreateNewGuid();
};

//Set Host web url in Local Storage
INTRANET.SetHostWebURlInSession = function() {
    var url = decodeURIComponent(INTRANET.GetQueryStringParameter("SPHostUrl"));
    localStorage.setItem("hostWebURL", url);
};

// Retrieve a query string value.  
INTRANET.GetQueryStringParameter = function(paramToRetrieve) {
    var params = document.URL.split("?")[1].split("&");
    for (var i = 0; i < params.length; i = i + 1) {
        var singleParam = params[i].split("=");
        if (singleParam[0] == paramToRetrieve) return singleParam[1];
    }
};

//Validation for Organization Name
INTRANET.CheckOrgName = function() {
    $("#lblError").css("display", "none");
    var orgName = $("#txtOrgName").val();

    if (orgName.trim() == "") {
        $("#lblError").css("display", "block");
    }
};

//Initial method for Creating INTRANET Site structure
INTRANET.CreateSiteStructure = function() {

    INTRANET.InIt();
    INTRANET.SetConfigSectionReadOnly();
    INTRANET.CreateSiteColumns().then(function() {
        INTRANET.CreateContentTypes().then(function() {
            //Add fields in COntent Type
            INTRANET.AddFieldsInContentType();
            addFieldsPromise.promise().then(function() {
                INTRANET.LoadAllContentTypes().then(function() {
                    INTRANET.AddContentTypeToList().then(function() {
                        INTRANET.SetAsDefaultContentType().then(function() {
                            INTRANET.AddFieldsDirectlyInList().then(function() {
                                INTRANET.BreakInheritedPermissions().then(function() {
                                    INTRANET.AddItemsInList().then(function() {
                                        INTRANET.InitiateMigration().then(function() {
                                            INTRANET.UploadMasterLookUpData().then(function() {
                                                INTRANET.UploadMasterPage();
                                                INTRANET.CreateAllFolders("Style Library", "IntranetInbox").then(function() {
                                                    var uploadUrl = INTRANET.GetRelativeUrlFromAbsolute(hostWebUrl) + "Style Library/IntranetInbox/",
                                                        readUrl = _spPageContextInfo.webAbsoluteUrl;
                                                    INTRANET.UploadAllFiles("Style Library", INTRANET.Schema.AllJS, uploadUrl + "js/", readUrl + "/IntranetJSFiles/");
                                                    INTRANET.UploadAllFiles("Style Library", INTRANET.Schema.AllFontFiles, uploadUrl + "fonts/", readUrl + "/fonts/");
                                                    INTRANET.UploadAllFiles("Style Library", INTRANET.Schema.AllImagesFiles, uploadUrl + "images/", readUrl + "/Images/");
                                                    INTRANET.UploadAllFiles("Style Library", INTRANET.Schema.AllCSS, uploadUrl + "css/", readUrl + "/IntranetCSSFiles/");
                                                    INTRANET.UploadAllFiles("Style Library", INTRANET.Schema.AllHtmlFiles, uploadUrl + "html/", readUrl + "/IntranetAllHtmlFiles/");
                                                    readUrl = _spPageContextInfo.webAbsoluteUrl + "/IntranetDisplayTemplateFiles/";
                                                    INTRANET.UploadAllDisplayTemplates(INTRANET.Schema.AllTemplateFiles, "_catalogs/masterpage/Display Templates/Content Web Parts", "html", readUrl)
                                                        .then(function() {
                                                            readUrl = _spPageContextInfo.webAbsoluteUrl + "/IntranetPageLayoutsFiles/";
                                                            INTRANET.UploadAllPageLayouts(INTRANET.Schema.AllPageLayouts, "_catalogs/masterpage", "html", readUrl).then(function() {
                                                                uploadUrl = INTRANET.GetRelativeUrlFromAbsolute(hostWebUrl) + "Pages/";
                                                                readUrl = _spPageContextInfo.webAbsoluteUrl + "/IntranetAllPagesFiles/";
                                                                INTRANET.UploadAllPageFiles("Pages", INTRANET.Schema.AllPages, uploadUrl, "aspx", readUrl).then(function() {
                                                                    readUrl = INTRANET.GetRelativeUrlFromAbsolute(hostWebUrl) + "Pages/Home.aspx";
                                                                    INTRANET.AddWebPartsToPage(readUrl);
                                                                    INTRANET.EndExecution();
                                                                });
                                                            });
                                                        });
                                                });
                                            });
                                        });
                                    });
                                });
                            });
                        });
                    });
                });
            });
        });
    });
};

//Execution for Look Up columns, content types and lists
INTRANET.CreateSiteStructureForLookUp = function() {
    var deferred = $.Deferred();
    countContentType = 0;
    INTRANET.App.IsLookUpExecution = true;
    $(".loader").show();
    INTRANET.CreateSiteColumns().then(function() {
        deferred.resolve();
    }, function() {
        deferred.resolve();
    });
    return deferred.promise();
};

//Create Site Columns on Host Web  
INTRANET.CreateSiteColumns = function() {
    var deferred = $.Deferred();
    try {
        INTRANET.Log(INTRANET.LogType.Info, "-----Creating Site Columns------");
        INTRANET.IncreaseProgressCounter("Configuring site columns... ");

        var asyncMethods = [];
        var tempSiteColumns = INTRANET.Schema.SiteColumns;

        if (INTRANET.App.IsLookUpExecution == true) {
            tempSiteColumns = INTRANET.Schema.SiteColumnsLookUp;
            INTRANET.Log(INTRANET.LogType.Info, "-----Creating LookUp Site Columns------");
        }
        if (tempSiteColumns) {
            asyncMethods = $.map(tempSiteColumns, function(pColumn) {

                if (INTRANET.App.IsLookUpExecution == true) {
                    pColumn = pColumn.XML;
                }

                INTRANET.Log(INTRANET.LogType.Info, "Site Column creating: " + pColumn);

                var dfd = $.Deferred();
                var fields;
                var createFields;
                //Get Host Web Fields Collection  
                fields = hostWeb.get_fields();

                //Parameters to pass  
                //fields.addFieldAsXml(string schemaXml,bool addToDefaultView, AddFieldOptions options)  
                createFields = fields.addFieldAsXml(pColumn, false, SP.AddFieldOptions.addFieldCheckDisplayName);

                context.load(fields);
                context.load(createFields);
                context.executeQueryAsync(
                    function(result) {
                        dfd.resolve();
                    },
                    function(sender, args) {
                        INTRANET.Log(INTRANET.LogType.Error, "Site Column Error: " + args.get_message());

                        var error = args.get_message();
                        var errorCode = args.get_errorCode();
                        //if (error.indexOf("A duplicate field name") >= 0) {
                        if (errorCode == -2146232832) {
                            dfd.resolve();
                        } else {
                            dfd.reject(args);
                        }
                    });
                return dfd.promise();
            });
        }
        //Below code will execute once all async call completed.
        $.when.apply($, asyncMethods).done(function(results) {
            console.log("-----Created Site Columns------");
            deferred.resolve();
            //Create Content Types

            //$.each(arguments, function (k, v) {
            //alert(v.get_count().toString());
            //})
        }).fail(function(e) {
            INTRANET.Log(INTRANET.LogType.Error, "Error in CreateSiteColumns. Error: " + e.get_message());
            deferred.reject();
        });
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in CreateSiteColumns. Error: " + ex.message);
    }
    return deferred.promise();
};

//Create Content Types on Host Web  
INTRANET.CreateContentTypes = function() {
    var deferred = $.Deferred();
    try {
        INTRANET.Log(INTRANET.LogType.Info, "-----Creating Content Types------");
        INTRANET.IncreaseProgressCounter("Configuring content types... ");

        var asyncMethods = [];

        asyncMethods = $.map(INTRANET.Schema.ContentTypes, function(objContentType) {

            //check host web context is not null and Current content type exist lookUp or Not
            if (typeof(hostWebContext) != 'undefined' && hostWebContext != null && objContentType.IsLookUpExist == INTRANET.App.IsLookUpExecution) {

                INTRANET.Log(INTRANET.LogType.Info, "Content Type creating: " + objContentType.Name);

                var dfd = $.Deferred();

                var hostWeb = hostWebContext.get_web();
                var contentTypeCollection = hostWeb.get_contentTypes();

                var contentType = contentTypeCollection.getById(objContentType.ParentContenttype);

                //Creating new Content Type  
                var newContentType = new SP.ContentTypeCreationInformation();
                newContentType.set_name(objContentType.Name);
                newContentType.set_group(INTRANET.Constant.ContentTypeGroupName);
                newContentType.set_description(objContentType.Description);

                //Set Base Content Type
                newContentType.set_parentContentType(contentType);
                contentTypeCollection.add(newContentType);

                context.load(contentTypeCollection);
                //context.load(newContentType);
                context.executeQueryAsync(
                    function() {
                        dfd.resolve();
                    },
                    function onContenttypeFailed(sender, args) {
                        INTRANET.Log(INTRANET.LogType.Info, "Content Type name. Error: " + args.get_message());

                        //var error = args.get_message();
                        var errorCode = args.get_errorCode();
                        //if (error.indexOf("A duplicate content type") >= 0) {
                        if (errorCode == 183) {
                            dfd.resolve();
                        } else {
                            dfd.reject(args);
                        }
                    });
                return dfd.promise();
            }
        });

        //Below code will execute once all async call completed.
        $.when.apply($, asyncMethods).done(function(results) {
            INTRANET.Log(INTRANET.LogType.Info, "-----Created Content Types------");
            deferred.resolve();

        }).fail(function(e) {
            INTRANET.Log(INTRANET.LogType.Error, "Error in CreateContentTypes. Error: " + e.get_message());
            deferred.reject();
        });
        return deferred.promise();
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in CreateContentTypes. Error: " + ex.message);
    }
};
//Add Fields in Content Type
INTRANET.AddFieldsInContentType = function() {

    try {

        INTRANET.Log(INTRANET.LogType.Info, "-----Add fields in Content Types------");

        ctColumns = [];
        if (typeof(hostWebContext) != 'undefined' && hostWebContext != null) {

            var objCurrentCT = INTRANET.Schema.ContentTypes[countContentType];

            if (objCurrentCT.IsLookUpExist == INTRANET.App.IsLookUpExecution) {

                var hostWeb = hostWebContext.get_web();
                //Get the columns that need to be added to Content type and store the objects in an Array.  
                for (var i = 0; i < objCurrentCT.FieldsInternalName.length; i++) {
                    ctColumns[i] = hostWeb.get_fields().getByInternalNameOrTitle(objCurrentCT.FieldsInternalName[i]);
                    context.load(ctColumns[i]);
                }

                //Get Host Web Content type Collection
                hostWebContentTypes = hostWeb.get_contentTypes();
                context.load(hostWebContentTypes);
                context.executeQueryAsync(
                    function() {

                        //Call function to add columns to Content Type  
                        INTRANET.CreateFieldsInContentType(INTRANET.Schema.ContentTypes[countContentType].Name, INTRANET.Schema.ContentTypes[countContentType].FieldsInternalName, ctColumns);

                    },
                    function onItemsRefetchedFail(sender, args) {
                        INTRANET.Log(INTRANET.LogType.Error, "Error in AddFieldsInContentType. Error: " + args.get_message() + '\n' + args.get_stackTrace());
                    });
            } else {
                countContentType++;
                if (countContentType == INTRANET.App.TotalSiteContentTypes) {
                    INTRANET.Log(INTRANET.LogType.Info, "-----Added fields in Content Types------");

                    //Call List Creation Method
                    INTRANET.CreateLists();
                } else {
                    //Call recursive function
                    INTRANET.AddFieldsInContentType();
                }
            }
        }
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in AddFieldsInContentType. Error: " + ex.message);
    }
};

// Add fields in existing Content Type
INTRANET.CreateFieldsInContentType = function(ctypeName, fieldsInternalName, createdFields) {
    try {
        //Find the Content Type and then add Fields to it  
        var createdContentType;
        var contentTypeEnumerator = hostWebContentTypes.getEnumerator();
        while (contentTypeEnumerator.moveNext()) {
            var contentType = contentTypeEnumerator.get_current();
            if (contentType.get_name() === ctypeName) {
                createdContentType = contentType;
                var fieldRef = [];
                for (var iAddFieldsCounter = 0; iAddFieldsCounter < createdFields.length; iAddFieldsCounter++) {
                    fieldRef[iAddFieldsCounter] = new SP.FieldLinkCreationInformation();
                    fieldRef[iAddFieldsCounter].set_field(createdFields[iAddFieldsCounter]);
                    createdContentType.get_fieldLinks().add(fieldRef[iAddFieldsCounter]);
                    createdContentType.update(true);
                }
                context.load(createdContentType);
                context.executeQueryAsync(
                    function() {
                        countContentType++;
                        if (countContentType == INTRANET.App.TotalSiteContentTypes) {
                            INTRANET.Log(INTRANET.LogType.Info, "-----Added fields in Content Types------");

                            //Call List Creation Method
                            INTRANET.CreateLists();
                        } else {
                            //Call recursive function
                            INTRANET.AddFieldsInContentType();
                        }
                    },
                    function onCreateFieldsInContentTypeFail(sender, args) {
                        INTRANET.Log(INTRANET.LogType.Error, "Error in onCreateFieldsInContentTypeFail. Error: " + args.get_message() + '\n' + args.get_stackTrace());
                        countContentType++;
                        if (countContentType == INTRANET.App.TotalSiteContentTypes) {
                            INTRANET.Log(INTRANET.LogType.Info, "-----Added fields in Content Types------");

                            //Call List Creation Method
                            INTRANET.CreateLists();
                        } else {
                            //Call recursive function
                            INTRANET.AddFieldsInContentType();
                        }
                    });
            }
        }
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in CreateFieldsInContentType. Error: " + ex.message);
    }
};

//Create Custom Lists
INTRANET.CreateLists = function() {
    try {

        INTRANET.Log(INTRANET.LogType.Info, "-----Creating Lists------");
        INTRANET.IncreaseProgressCounter("Configuring lists... ");

        INTRANET.App.CollLists = [];

        var asyncMethods = [];
        var counter = 0;
        asyncMethods = $.map(INTRANET.Schema.Lists, function(objList) {

            if (typeof(hostWebContext) != 'undefined' && hostWebContext != null && objList.IsLookUpExist == INTRANET.App.IsLookUpExecution) {

                INTRANET.Log(INTRANET.LogType.Info, "List creating: " + objList.Name);

                var dfd = $.Deferred();

                var hostWeb = hostWebContext.get_web();
                var listcreation = new SP.ListCreationInformation();
                listcreation.set_title(objList.Name);
                listcreation.set_templateType(SP.ListTemplateType[objList.Type]);
                //listcreation.set_contentTypesEnabled(true);
                //var list = hostWeb.get_lists().add(listcreation);
                INTRANET.App.CollLists[counter] = hostWeb.get_lists().add(listcreation);
                //var lists = hostWebContext.get_web().get_lists();
                //var list = lists.add(listcreation);
                //                //Hide List
                //if (objList.IsHidden != null && typeof (objList.IsHidden) != 'undefined' && objList.IsHidden == true) {
                //    list.set_hidden(hidden);
                //    list.update();
                //}
                //INTRANET.App.CollLists[counter] = list;

                //context.load(list);
                context.load(INTRANET.App.CollLists[counter]);
                counter++;
                context.executeQueryAsync(
                    function() {
                        dfd.resolve();
                    },
                    function onCreateListFailed(sender, args) {

                        INTRANET.Log(INTRANET.LogType.Info, "onCreateListFailed. Error: " + args.get_message());

                        //var error = args.get_message();
                        //if (error.indexOf("Please choose another title") >= 0) {
                        var errorCode = args.get_errorCode();
                        if (errorCode == -2130575342) {
                            dfd.resolve();
                        } else {
                            dfd.reject(args);
                        }
                    });
                return dfd.promise();
            }
        });

        //Below code will execute once all async call completed.
        $.when.apply($, asyncMethods).done(function(results) {
            INTRANET.Log(INTRANET.LogType.Info, "-----Created Lists------");

            if (INTRANET.App.IsLookUpExecution == false) {
                INTRANET.SetListGUID();
            }

            addFieldsPromise.resolve();

        }).fail(function(e) {
            INTRANET.Log(INTRANET.LogType.Error, "Error in CreateLists. Error: " + e.get_message());
            addFieldsPromise.resolve();
        });
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in CreateLists. Error: " + ex.message);
    }
    return addFieldsPromise.promise();
};

INTRANET.SetListGUID = function() {
    try {
        for (var i = 0; i < INTRANET.App.CollLists.length; i++) {
            var listName = '';
            var listId = '';

            try {
                listName = INTRANET.App.CollLists[i].get_title();
                listId = INTRANET.App.CollLists[i].get_id();
                listId = listId.toString();
            } catch (ex1) {}

            if (INTRANET.Schema.SiteColumnsLookUp && listId != '' && listId != null && typeof(listId) != 'undefined') {

                //set master list id in site lookup column collection
                for (var j = 0; j < INTRANET.Schema.SiteColumnsLookUp.length; j++) {
                    if (INTRANET.Schema.SiteColumnsLookUp[j].MasterListName == listName) {
                        INTRANET.Schema.SiteColumnsLookUp[j].MasterListId = listId;

                        var tempXML = INTRANET.Schema.SiteColumnsLookUp[j].XML;

                        INTRANET.Schema.SiteColumnsLookUp[j].XML = tempXML.replace("$ListName$", listId);
                    }
                }
            }
        }
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in SetListGUID. Error: " + ex.message);
    }
};

//Load All Content Types
INTRANET.LoadAllContentTypes = function() {
    var deferred = $.Deferred();
    try {
        INTRANET.IncreaseProgressCounter("Configuring load content types... ");

        var hostWeb = hostWebContext.get_web();

        var allContentTypeColl = hostWeb.get_contentTypes();
        //load the web , lists & content types  
        context.load(allContentTypeColl);
        context.executeQueryAsync(function() {
                var contentTypeEnum = allContentTypeColl.getEnumerator();
                while (contentTypeEnum.moveNext()) {

                    var currentCT = contentTypeEnum.get_current();
                    $.each(INTRANET.Schema.Lists, function(key, objListSchema) {
                        if (currentCT.get_name() == objListSchema.ContentType) {
                            var ctID = currentCT.get_stringId();
                            objListSchema.ContentTypeID = ctID;
                        }
                    });
                }
                deferred.resolve();
            },
            function(sender, args) {
                console.error("Error in LoadAllContentTypes. Error: " + e.get_message());
            });
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in LoadAllContentTypes. Error: " + ex.message);
    }
    return deferred.promise();
};

//Attach Content Types to List
INTRANET.AddContentTypeToList = function() {
    var deferred = $.Deferred();
    try {

        INTRANET.Log(INTRANET.LogType.Info, "-----Attaching content type to Lists------");

        var asyncMethods = [];

        asyncMethods = $.map(INTRANET.Schema.Lists, function(objList) {

            if (typeof(hostWebContext) != 'undefined' && hostWebContext != null && objList.IsLookUpExist == INTRANET.App.IsLookUpExecution) {
                if (typeof(objList.ContentTypeID) != 'undefined' && objList.ContentTypeID != null && objList.ContentTypeID != '') {

                    INTRANET.Log(INTRANET.LogType.Info, "List processing: " + objList.Name);

                    var dfd = $.Deferred();


                    var hostWeb = hostWebContext.get_web();
                    var contentTypeCol = hostWeb.get_contentTypes();
                    var contentType = contentTypeCol.getById(objList.ContentTypeID);
                    var targetList = hostWeb.get_lists().getByTitle(objList.Name);
                    targetList.set_contentTypesEnabled(true);
                    var cts = targetList.get_contentTypes();
                    cts.addExistingContentType(contentType);
                    targetList.update();
                    context.load(cts);
                    context.executeQueryAsync(
                        function() {

                            dfd.resolve();
                        },
                        function(sender, args) {

                            //var error = args.get_message();
                            //if (error.indexOf("A duplicate content type") >= 0) {
                            var errorCode = args.get_errorCode();
                            if (errorCode == 183) {
                                dfd.resolve();
                            } else {
                                dfd.reject(args);
                            }
                            dfd.reject(args);
                        });
                    return dfd.promise();
                }
            }
        });

        //Below code will execute once all async call completed.
        $.when.apply($, asyncMethods).done(function(results) {
            INTRANET.Log(INTRANET.LogType.Info, "-----Attached content type to Lists------");
            deferred.resolve();
        }).fail(function(e) {
            INTRANET.Log(INTRANET.LogType.Error, "Error in AddContentTypeToList. Error: " + e.get_message());
            deferred.resolve();
        });

    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in AddContentTypeToList. Error: " + ex.message);
    }

    return deferred.resolve();
};

//Change default contend type (this code will reverse the current content type order, so our recently added content type will come first)
INTRANET.SetAsDefaultContentType = function() {
    var deferred = $.Deferred();
    try {

        INTRANET.Log(INTRANET.LogType.Info, "-----Attaching content type to Lists------");

        var asyncMethods = [];

        asyncMethods = $.map(INTRANET.Schema.Lists, function(objList) {

            if (typeof(hostWebContext) != 'undefined' && hostWebContext != null && objList.SetAsDefaultContentType == true && objList.IsLookUpExist == INTRANET.App.IsLookUpExecution) {

                INTRANET.Log(INTRANET.LogType.Info, "List processing for default content type: " + objList.Name);

                var dfd = $.Deferred();

                var hostWeb = hostWebContext.get_web();
                var folder = hostWeb.get_lists().getByTitle(objList.Name).get_rootFolder();
                var contentTypeCollection = hostWeb.get_lists().getByTitle(objList.Name).get_contentTypes();
                context.load(folder, 'ContentTypeOrder');
                context.load(contentTypeCollection);
                context.executeQueryAsync(
                    function() {

                        var contentTypeEnumerator = contentTypeCollection.getEnumerator();
                        var conteTypeID = '';
                        var flag = false;

                        while (contentTypeEnumerator.moveNext()) {
                            var content = contentTypeEnumerator.get_current();
                            var contentTypeName = content.get_name();

                            //check is content type exist in our collection
                            var tempColl = $.grep(INTRANET.Schema.Lists, function(n, i) {
                                return n["ContentType"] == contentTypeName;
                            });

                            if (tempColl.length > 0) {
                                conteTypeID = content.get_id().toString();
                            }
                        }

                        if (conteTypeID == folder.get_contentTypeOrder()[0].toString()) {
                            flag = true;
                        }

                        //check if current content type is first content type in list
                        if (flag == false) {
                            folder.set_uniqueContentTypeOrder(folder.get_contentTypeOrder().reverse());
                            folder.update();
                            context.executeQueryAsync(function onSuccess() {
                                    INTRANET.Log(INTRANET.LogType.Info, 'Content type changed!!!');
                                },
                                function onFail() {
                                    INTRANET.Log(INTRANET.LogType.Info, 'Error to set on default content type');
                                });
                        }

                        dfd.resolve();
                    },
                    function(sender, args) {

                        dfd.reject(args);
                    });
                return dfd.promise();
            }
        });

        //Below code will execute once all async call completed.
        $.when.apply($, asyncMethods).done(function(results) {
            INTRANET.Log(INTRANET.LogType.Info, "-----Default content type set to all Lists------");

            deferred.resolve();

        }).fail(function(e) {
            INTRANET.Log(INTRANET.LogType.Error, "Error in SetAsDefaultContentType. Error: " + e.get_message());
        });

    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in SetAsDefaultContentType. Error: " + ex.message);
    }

    return deferred.promise();
};

//Add column directly in List - List level columns
INTRANET.AddFieldsDirectlyInList = function() {
    var deferred = $.Deferred();
    try {

        INTRANET.Log(INTRANET.LogType.Info, "-----Adding List level columns------");

        var asyncMethodsList = [];
        var asyncMethods = [];

        asyncMethodsList = $.map(INTRANET.Schema.Lists, function(objList) {
            if (typeof(hostWebContext) != 'undefined' && hostWebContext != null && objList.IsLookUpExist == INTRANET.App.IsLookUpExecution) {

                console.log("List level column processing for List: " + objList.Name);
                var dfd = $.Deferred();

                var hostWeb = hostWebContext.get_web();
                var oList = hostWeb.get_lists().getByTitle(objList.Name);
                asyncMethods = $.map(objList.Fields, function(objListField) {

                    var newField = oList.get_fields().addFieldAsXml(objListField, true, SP.AddFieldOptions.addToDefaultContentType);
                    newField.update();

                });
                context.executeQueryAsync(
                    function() {

                        dfd.resolve();
                    },
                    function(sender, args) {

                        dfd.reject(args);
                    });
                return dfd.promise();
            }
        });

        //Below code will execute once all async call completed.
        $.when.apply($, asyncMethodsList).done(function(results) {
            INTRANET.Log(INTRANET.LogType.Info, "-----Added List level columns------");
            INTRANET.Log(INTRANET.LogType.Info, "-----END Execution for AddFieldsDirectlyInList------");

            $(".loader").hide();

            if (INTRANET.App.IsLookUpExecution == false) {
                INTRANET.CreateSiteStructureForLookUp();
            } else {
                INTRANET.Log(INTRANET.LogType.Info, "-----Added List level columns for Look Up------");
                INTRANET.Log(INTRANET.LogType.Info, "-----END Execution for Look Up------");
            }
            deferred.resolve();

        }).fail(function(e) {
            INTRANET.Log(INTRANET.LogType.Error, "Error in AddFieldsDirectlyInList. Error: " + e.get_message());
        });

    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in AddFieldsDirectlyInList. Error: " + ex.message);
    }

    return deferred.promise();
};

//Break Inherited permissions of Lists
INTRANET.BreakInheritedPermissions = function() {
    var deferred = $.Deferred();
    try {
        $(".loader").show();

        INTRANET.Log(INTRANET.LogType.Info, "-----Started Break List Permissions------");

        var asyncMethods = [];

        asyncMethods = $.map(INTRANET.Schema.Lists, function(objList) {
            if (typeof(hostWebContext) != 'undefined' && hostWebContext != null && objList.IsUniquePermission == true) {

                console.log("Inheritance Break started for List: " + objList.Name);
                var dfd = $.Deferred();

                var hostWeb = hostWebContext.get_web();

                var oList = hostWeb.get_lists().getByTitle(objList.Name);
                oList.breakRoleInheritance(false, true);

                // All users , EveryOne , All Athenticated Users
                var userobj = hostWeb.ensureUser("c:0(.s|true");
                var role = SP.RoleDefinitionBindingCollection.newObject(context);
                role.add(hostWeb.get_roleDefinitions().getByType(SP.RoleType.reader));
                oList.get_roleAssignments().add(userobj, role);

                context.executeQueryAsync(
                    function() {

                        dfd.resolve();
                    },
                    function(sender, args) {

                        dfd.reject(args);
                    });
                return dfd.promise();
            }
        });

        //Below code will execute once all async call completed.
        $.when.apply($, asyncMethods).done(function(results) {
            INTRANET.Log(INTRANET.LogType.Info, "-----Done Break List Permissions------");
            INTRANET.Log(INTRANET.LogType.Info, "-----END Execution for Break List Permissions------");

            deferred.resolve();
        }).fail(function(e) {
            INTRANET.Log(INTRANET.LogType.Error, "Error in BreakInheritedPermissions. Error: " + e.get_message());
        });

    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in BreakInheritedPermissions. Error: " + ex.message);
    }

    return deferred.promise()
};

//Add Default Items in List
INTRANET.AddItemsInList = function() {
    var deferred = $.Deferred();
    try {
        $(".loader").show();
        INTRANET.IncreaseProgressCounter("Configuring default config items... ");

        INTRANET.Log(INTRANET.LogType.Info, "-----Adding Items in list------");

        var asyncMethodsList = [];
        var asyncMethods = [];

        asyncMethodsList = $.map(INTRANET.Schema.Lists, function(objList) {
            if (typeof(hostWebContext) != 'undefined' && hostWebContext != null) {
                if (typeof(objList.Items) != 'undefined' && objList.Items != null && objList.Items.length > 0) {

                    console.log("Item adding for List: " + objList.Name);
                    var dfd = $.Deferred();

                    var hostWeb = hostWebContext.get_web();
                    var oList = hostWeb.get_lists().getByTitle(objList.Name);
                    asyncMethods = $.map(objList.Items, function(fields) {
                        var itemCreateInfo = new SP.ListItemCreationInformation();
                        var oListItem = oList.addItem(itemCreateInfo);
                        var dynamicValue = '';

                        $.map(fields, function(field) {
                            oListItem.set_item(field.Key, field.Value);
                        });

                        oListItem.update();
                    });
                    context.executeQueryAsync(
                        function() {
                            dfd.resolve();
                        },
                        function(sender, args) {
                            INTRANET.Log(INTRANET.LogType.Info, "AddItemsInListFailed. Error: " + args.get_message());

                            //var error = args.get_message();
                            //if (error.indexOf("duplicate values were found") >= 0) {
                            var errorCode = args.get_errorCode();
                            if (errorCode == -2130575169) {
                                dfd.resolve();
                            } else {
                                dfd.reject(args);
                            }
                            dfd.reject(args);
                        });
                    return dfd.promise();
                }
            }
        });

        //Below code will execute once all async call completed.
        $.when.apply($, asyncMethodsList).done(function(results) {
            INTRANET.Log(INTRANET.LogType.Info, "-----Items added to the list------");
            INTRANET.Log(INTRANET.LogType.Info, "-----END Execution of Site Structure creation------");

            $(".loader").hide();
            //$("#configureRead").css("display", "block");

            //Initiate Migration
            deferred.resolve();
        }).fail(function(e) {
            INTRANET.Log(INTRANET.LogType.Error, "Error in AddItemsInList. Error: " + e.get_message());
        });
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in AddItemsInList. Error: " + ex.message);
    }
    return deferred.promise();
};

INTRANET.SetConfigSectionReadOnly = function() {
    $("#txtOrgName").prop('disabled', true);
    $("#configureBtn").prop('disabled', true);
    $("#txtShortAddress").prop('disabled', true);
    $("#configureBtn").css('opacity', "0.65");
};

//Start Migration activity - Upload data in master Lists
INTRANET.InitiateMigration = function() {
    var deferred = $.Deferred();
    try {
        INTRANET.UpdateInstallationState(INTRANET.App.APPState.Initialization, INTRANET.App.Done);
        //Set list as Hidden
        INTRANET.SetListToHidden();

        INTRANET.Log(INTRANET.LogType.Info, "----Started Migration----");
        INTRANET.IncreaseProgressCounter("Initialized sample data upload... ");

        hostWebContextFile = new SP.ClientContext(INTRANET.GetRelativeUrlFromAbsolute(hostWebUrl));
        var web = hostWebContextFile.get_web();
        hostWebContextFile.load(web);
        hostWebContextFile.executeQueryAsync(
            function() {
                INTRANET.Log(INTRANET.LogType.Info, "Context Created");
                deferred.resolve();
            },
            function(sender, args) {
                INTRANET.Log(INTRANET.LogType.Error, "Error in InitiateMigration AJAX. Error: " + args.get_errorDetails());
            })
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in InitiateMigration. Error: " + ex.message);
    }
    return deferred.promise();
};

//Migrate Look Up column contains Master List data
INTRANET.UploadMasterLookUpData = function() {
    var deferred = $.Deferred();
    try {
        INTRANET.IncreaseProgressCounter("Configuring sample data...");

        INTRANET.Log(INTRANET.LogType.Info, "-----Master List data look Up uploading------");

        var asyncMethodsList = [];
        var asyncMethods = [];
        if (INTRANET.Schema.MasterDataLookUp) {
            asyncMethodsList = $.map(INTRANET.Schema.MasterDataLookUp, function(objList) {

                if (typeof(hostWebContext) != 'undefined' && hostWebContext != null) {
                    if (typeof(objList.Items) != 'undefined' && objList.Items != null && objList.Items.length > 0) {

                        INTRANET.Log(INTRANET.LogType.Info, "Item adding for Master List Look Up: " + objList.ListName);

                        var dfdParent = $.Deferred();

                        asyncMethods = $.map(objList.Items, function(objListItem) {
                            var dfd = $.Deferred();


                            var sourcePath = "/Images/" + objList.ListName + "/" + objListItem.FileName;
                            var targetPath = "/" + objList.ListName + "/" + objListItem.FileName;

                            $.ajax({
                                url: INTRANET.Constant.AppWebURL + sourcePath,
                                type: "GET",
                                dataType: "binary",
                                processData: false,
                                responseType: 'arraybuffer',
                                cache: false,
                                async: false
                            }).done(function(contents) {

                                INTRANET.Log(INTRANET.LogType.Info, "Image fetched: " + sourcePath);

                                var fileName = INTRANET.GetFilenameFromUrl(targetPath);
                                var folder = INTRANET.GetPathFromUrl(targetPath);

                                var createInfo = new SP.FileCreationInformation();
                                // Convert ArrayBuffer to Base64 string
                                createInfo.set_content(INTRANET.ArrayBufferToBase64(contents));
                                // Overwrite if already exists
                                createInfo.set_overwrite(true);
                                // set target url
                                createInfo.set_url(fileName);
                                // retrieve file collection of folder
                                var files = hostWebContextFile.get_web().getFolderByServerRelativeUrl(INTRANET.GetRelativeUrlFromAbsolute(hostWebUrl) + folder).get_files();
                                // load file collection from host web
                                hostWebContextFile.load(files);
                                var newFile = files.add(createInfo);

                                //file MetaData
                                var objLstItem = newFile.get_listItemAllFields();

                                //set item properties
                                for (var c = 0; c < objListItem.Columns.length; c++) {
                                    objLstItem.set_item(objListItem.Columns[c]["FieldName"], objListItem.Columns[c]["FieldValue"]);
                                }

                                //update item
                                objLstItem.update();
                                hostWebContextFile.load(newFile);

                                // upload file
                                hostWebContextFile.executeQueryAsync(function(response) {

                                    INTRANET.Log(INTRANET.LogType.Info, "File Uploaded in Master List Look Up: " + response);
                                    dfd.resolve(response);
                                }, function(sender, args) {

                                    INTRANET.Log(INTRANET.LogType.Info, "UploadMasterLookUpData AJAX File Upload Look Up. Error: " + args.get_message());
                                    dfd.reject(args);
                                });
                            }).fail(function(jqXHR, textStatus) {

                                INTRANET.Log(INTRANET.LogType.Error, "Error in UploadMasterLookUpData Ajax call Look Up: " + textStatus);
                                dfd.reject();
                            });

                            return dfd.promise();
                        });

                        //Below code will execute once all async call completed.
                        $.when.apply($, asyncMethods).done(function(results) {

                            dfdParent.resolve(results);

                            INTRANET.Log(INTRANET.LogType.Info, "-----Master data uploading look Up...------");
                        }).fail(function(e) {

                            INTRANET.Log(INTRANET.LogType.Error, "Error in UploadMasterLookUpData when apply child look up. Error: " + e.get_message());
                            dfd.reject(e.get_message());
                        });

                        return dfdParent.promise();
                    }
                }
            });
        }
        //Below code will execute once all async call completed.
        $.when.apply($, asyncMethodsList).done(function(results) {

            countTotalMasterLookUpLists++;
            INTRANET.Log(INTRANET.LogType.Info, "-----Master data uploaded Look Up------");
            $(".loader").hide();

            //once all items processed execute next code
            if (countTotalMasterLookUpLists == INTRANET.App.TotalMasterLookUpLists) {
                INTRANET.LoadAllCategories();
            } else {
                deferred.resolve();
            }

        }).fail(function(e) {

            countTotalMasterLookUpLists++;
            INTRANET.Log(INTRANET.LogType.Error, "Error in UploadMasterLookUpData when apply. Error: " + e.get_message());
            //once all items processed execute next code

            if (countTotalMasterLookUpLists == INTRANET.App.TotalMasterLookUpLists) {
                INTRANET.LoadAllCategories();
            } else {
                INTRANET.UploadMasterPage();
                INTRANET.CreateAllFolders("Style Library", "NBB/CSS");
                var uploadUrl = INTRANET.GetRelativeUrlFromAbsolute(hostWebUrl) + "Style Library/NBB/",
                    readUrl = _spPageContextInfo.webAbsoluteUrl + "/IntranetJSFiles/";
                INTRANET.UploadAllFiles(INTRANET.Schema.AllJS, uploadUrl, "js", readUrl);
                readUrl = _spPageContextInfo.webAbsoluteUrl + "/IntranetCSSFiles/";
                INTRANET.UploadAllFiles(INTRANET.Schema.AllCSS, uploadUrl, "css", readUrl);
                readUrl = _spPageContextInfo.webAbsoluteUrl + "/IntranetTextFiles/";
                INTRANET.UploadAllFiles(INTRANET.Schema.AllTextFiles, uploadUrl, "txt", readUrl);
                INTRANET.EndExecution();
            }
        });
    } catch (ex) {

        INTRANET.Log(INTRANET.LogType.Error, "Error in UploadMasterLookUpData. Error: " + ex.message);
    }
    return deferred.promise();
};

//Load All Categories
INTRANET.LoadAllCategories = function() {

    try {

        var hostWeb = hostWebContext.get_web();

        var oList = hostWeb.get_lists().getByTitle('Categories');
        var camlQuery = new SP.CamlQuery();
        camlQuery.set_viewXml('<View></View>');
        var allCategoriesColl = oList.getItems(camlQuery);
        context.load(allCategoriesColl);
        context.executeQueryAsync(function() {
                var categoryEnum = allCategoriesColl.getEnumerator();
                while (categoryEnum.moveNext()) {

                    var lstItem = categoryEnum.get_current();
                    $.each(INTRANET.Schema.MasterDataLookUp, function(key, objListSchema) {
                        if (objListSchema.ListName == "Categories") {
                            $.each(objListSchema.Items, function(key, objItem) {

                                if (lstItem.get_item("CategoryName") == objItem.Category) {
                                    var ctID = lstItem.get_id();
                                    objItem.ID = ctID.toString();
                                }
                            });
                        }
                    });
                }

                //Upload Master List Data
                INTRANET.UploadMasterData();
            },
            function(sender, args) {
                INTRANET.Log(INTRANET.LogType.Error, "Error in LoadAllCategories. Error: " + args);
            });
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in LoadAllCategories. Error: " + ex.message);
    }
};

INTRANET.UploadMasterData = function() {
    try {

        $(".loader").show();
        INTRANET.IncreaseProgressCounter("Configuring sample data...");

        INTRANET.Log(INTRANET.LogType.Info, "-----Master List data uploading------");

        var asyncMethodsList = [];
        var asyncMethods = [];

        asyncMethodsList = $.map(INTRANET.Schema.MasterData, function(objList) {

            if (typeof(hostWebContext) != 'undefined' && hostWebContext != null) {
                if (typeof(objList.Items) != 'undefined' && objList.Items != null && objList.Items.length > 0) {

                    INTRANET.Log(INTRANET.LogType.Info, "Item adding for Master List: " + objList.ListName);

                    var dfdParent = $.Deferred();

                    asyncMethods = $.map(objList.Items, function(objListItem) {
                        var dfd = $.Deferred();


                        var sourcePath = "/Images/" + objList.ListName + "/" + objListItem.FileName;
                        var targetPath = "/" + objList.ListName + "/" + objListItem.FileName;
                        var listName = objList.ListName;

                        $.ajax({
                            url: INTRANET.Constant.AppWebURL + sourcePath,
                            type: "GET",
                            dataType: "binary",
                            processData: false,
                            responseType: 'arraybuffer',
                            cache: false,
                            async: false
                        }).done(function(contents) {

                            INTRANET.Log(INTRANET.LogType.Info, "Image fetched: " + sourcePath);

                            var fileName = INTRANET.GetFilenameFromUrl(targetPath);
                            var folder = INTRANET.GetPathFromUrl(targetPath);

                            var createInfo = new SP.FileCreationInformation();
                            // Convert ArrayBuffer to Base64 string
                            createInfo.set_content(INTRANET.ArrayBufferToBase64(contents));
                            // Overwrite if already exists
                            createInfo.set_overwrite(true);
                            // set target url
                            createInfo.set_url(fileName);
                            // retrieve file collection of folder
                            var files = hostWebContextFile.get_web().getFolderByServerRelativeUrl(INTRANET.GetRelativeUrlFromAbsolute(hostWebUrl) + folder).get_files();
                            // load file collection from host web
                            hostWebContextFile.load(files);
                            var newFile = files.add(createInfo);

                            //file MetaData
                            var objLstItem = newFile.get_listItemAllFields();

                            //set item properties
                            for (var c = 0; c < objListItem.Columns.length; c++) {
                                if (objListItem.Columns[c]["IsLookUp"] != true) {
                                    objLstItem.set_item(objListItem.Columns[c]["FieldName"], objListItem.Columns[c]["FieldValue"]);
                                } else {
                                    //get list item ID
                                    var itemID = INTRANET.GetListitemID(objListItem.Category, "Categories");

                                    //Lookup Column
                                    var lookUpValue = new SP.FieldLookupValue();
                                    lookUpValue.set_lookupId(itemID);
                                    objLstItem.set_item(objListItem.Columns[c]["FieldName"], lookUpValue);
                                }
                            }

                            //update item
                            objLstItem.update();
                            hostWebContextFile.load(newFile);

                            // upload file
                            hostWebContextFile.executeQueryAsync(function(response) {

                                INTRANET.Log(INTRANET.LogType.Info, "File Uploaded in Master List: " + response);
                                dfd.resolve(response);
                            }, function(sender, args) {

                                INTRANET.Log(INTRANET.LogType.Info, "ReadFileFromAppWeb AJAX File Upload. Error: " + args.get_message());
                                dfd.reject(args);
                            });
                        }).fail(function(jqXHR, textStatus) {

                            INTRANET.Log(INTRANET.LogType.Error, "Error in ReadFileFromAppWeb Ajax call: " + textStatus);
                            dfd.resolve(textStatus);
                        });

                        return dfd.promise();
                    });

                    //Below code will execute once all async call completed.
                    $.when.apply($, asyncMethods).done(function(results) {

                        INTRANET.Log(INTRANET.LogType.Info, "-----Master data uploading...------");

                        dfdParent.resolve(results);

                    }).fail(function(e) {

                        INTRANET.Log(INTRANET.LogType.Error, "Error in ReadFileFromAppWeb when apply child. Error: " + e.get_message());
                    });

                    return dfdParent.promise();
                }
            }
        });

        //Below code will execute once all async call completed.
        $.when.apply($, asyncMethodsList).done(function(results) {
            INTRANET.Log(INTRANET.LogType.Info, "-----Master data uploaded------");

            //INTRANET.CreateSharePointGroup();
            //END Execution

            INTRANET.EndExecution();

            //INTRANET.UploadSampleData();

        }).fail(function(e) {

            INTRANET.Log(INTRANET.LogType.Error, "Error in ReadFileFromAppWeb when apply. Error: " + e.get_message());
        });
    } catch (ex) {

        INTRANET.Log(INTRANET.LogType.Error, "Error in UploadMasterData. Error: " + ex.message);
    }
};

//Upload some items in master List from Azure
INTRANET.UploadSampleData = function() {
    try {
        INTRANET.Log(INTRANET.LogType.Info, "-----Uploading sample data------");

        $.ajax({
            url: INTRANET.Constant.WebAPIURL + INTRANET.App.UploadSampleDataAPI,
            type: "POST",
            data: JSON.stringify({
                TenantURL: INTRANET.Constant.HostWebURL,
                TenantId: INTRANET.Constant.TenantId
            }),
            datatype: "json",
            contentType: 'application/json',
            async: false,
            cache: false,
            success: function(data, status, xhr) {
                INTRANET.Log(INTRANET.LogType.Info, "-----Sample data uploaded------");
                //END Execution
                INTRANET.EndExecution();
            },
            error: function(xhr, status, error) {
                INTRANET.Log(INTRANET.LogType.Error, "Error in ajax UploadSampleData. Error: " + xhr.responseText);
            }
        });
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in UploadSampleData. Error: " + ex.message);
    }
};

//Fetch all cards from azure and upload on sharepoint
INTRANET.GetAllCards = function() {
    try {
        INTRANET.Log(INTRANET.LogType.Info, "-----Fetching all cards------");

        $("#idRollBackConfirmation").show();

        $.ajax({
            url: INTRANET.Constant.WebAPIURL + INTRANET.App.GetAllCardsAPI,
            type: "POST",
            data: JSON.stringify({
                TenantURL: INTRANET.Constant.HostWebURL,
                TenantId: INTRANET.Constant.TenantId
            }),
            datatype: "json",
            contentType: 'application/json',
            async: true,
            cache: false,
            success: function(data, status, xhr) {
                //INTRANET.Log(INTRANET.LogType.Info, "-----Sample data uploaded------");                
            },
            error: function(xhr, status, error) {
                INTRANET.Log(INTRANET.LogType.Error, "Error in ajax GetAllCards. Error: " + xhr.responseText);
                $("#idRollBackConfirmation").show();
                $("#idRollBackConfirmation").innerText = "Something went wrong. Please contact to support team.";
            }
        });
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in GetAllCards. Error: " + ex.message);
    }
};

//End Execution
INTRANET.EndExecution = function() {
    try {
        INTRANET.IncreaseProgressCounter("Configuration successfully completed... ");
        INTRANET.UpdateInstallationState(INTRANET.App.APPState.SiteConfiguration, INTRANET.App.Done);

        INTRANET.Log(INTRANET.LogType.Info, "-----END Execution------");

        $(".loader").hide();
        $("#configureRead").css("display", "block");
        $("#idInstallStatusText").css("display", "none");
        $("#idInstallStatusCount").css("display", "none");
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in EndExecution. Error: " + ex.message);
    }
};

//Fetch List Item id from collection as per category name
INTRANET.GetListitemID = function(strCategory, steListName) {
    var itemID = 0;
    for (var c = 0; c < INTRANET.Schema.MasterDataLookUp.length; c++) {
        if (INTRANET.Schema.MasterDataLookUp[c].ListName == steListName) {
            var items = INTRANET.Schema.MasterDataLookUp[c].Items;
            for (var i = 0; i < items.length; i++) {
                if (items[i].Category == strCategory) {
                    itemID = items[i].ID;
                    break;
                }
            }
        }
    }

    return itemID;
};

//Installation Progress
INTRANET.IncreaseProgressCounter = function(pStatusText) {
    try {
        counterInstallation = counterInstallation + 1;
        if (INTRANET.App.IsRollback == true) {
            $("#idInstallRollbackText").html(pStatusText)
            $("#idRollBackCount").html(counterInstallation);
        } else {
            $("#idInstallStatusText").html(pStatusText)
            $("#idInstallCount").html(counterInstallation);
        }
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in IncreaseProgressCounter. Error: " + ex.message);
    }
};

//Update Installation State of APP
INTRANET.UpdateInstallationState = function(pKeyword, pCurrentState) {
    try {
        if (typeof(hostWebContext) != 'undefined' && hostWebContext != null) {

            var hostWeb = hostWebContext.get_web();
            var oList = hostWeb.get_lists().getByTitle(INTRANET.App.ListBravoAppState);

            var itemCreateInfo = new SP.ListItemCreationInformation();
            var oListItem = oList.addItem(itemCreateInfo);
            oListItem.set_item('BravoKey', pKeyword);
            oListItem.set_item('BravoValue', pCurrentState);

            oListItem.update();
            context.executeQueryAsync(
                function() {
                    INTRANET.Log(INTRANET.LogType.Info, "Installation Status Updated in list.");
                },
                function(sender, args) {
                    //INTRANET.Log(INTRANET.LogType.Error, "UpdateInstallationStateFailed. Error: " + args.get_message());
                });
        }
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in UpdateInstallationState. Error: " + ex.message);
    }
};

//Create SahrePoint Security Group
//this method is used to create a group  
INTRANET.CreateSharePointGroup = function() {
    try {
        INTRANET.Log(INTRANET.LogType.Info, "-----Creating security groups------");

        var hostWeb = hostWebContext.get_web();

        // Create Group information for Group  
        var membersGRP = new SP.GroupCreationInformation();
        membersGRP.set_title(INTRANET.App.SPGroup.Name);
        membersGRP.set_description(INTRANET.App.SPGroup.Description);
        //add group  
        var oMembersGRP = hostWeb.get_siteGroups().add(membersGRP);
        //return SP.RoleDefinition object  
        var rdContribute = hostWeb.get_roleDefinitions().getByName("Read");
        // Create a new RoleDefinitionBindingCollection.  
        var collContribute = SP.RoleDefinitionBindingCollection.newObject(context);
        // Add the role to the collection.  
        collContribute.add(rdContribute);
        // Get the RoleAssignmentCollection for the target web.  
        var assignments = hostWeb.get_roleAssignments();
        // assign the group to the new RoleDefinitionBindingCollection.  
        var roleAssignmentContribute = assignments.add(oMembersGRP, collContribute);
        context.load(oMembersGRP);
        context.executeQueryAsync(function() {
            INTRANET.Log(INTRANET.LogType.Info, "-----Created security groups------");
            INTRANET.InsertCurrentUserSPGroup();

            //END Execution
            INTRANET.EndExecution();

        }, function(sender, args) {
            INTRANET.Log(INTRANET.LogType.Error, "Error in CreateSharePointGroup execute Query. Error: " + args);

            var errorCode = args.get_errorCode();
            if (errorCode == -2130575293) {
                INTRANET.InsertCurrentUserSPGroup();

                //END Execution
                INTRANET.EndExecution();
            }
        });
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in CreateSharePointGroup. Error: " + ex.message);
    }
};

//Create SahrePoint Security Group
//this method is used to create a group  
INTRANET.InsertCurrentUserSPGroup = function() {
    try {
        INTRANET.Log(INTRANET.LogType.Info, "-----Inserting current user in admin group------");

        var hostWeb = hostWebContext.get_web();
        var siteGroups = hostWeb.get_siteGroups();

        var spGroup = siteGroups.getByName(INTRANET.App.SPGroup.Name);
        var user = hostWeb.ensureUser(_spPageContextInfo.userLoginName);
        var userCollection = spGroup.get_users();
        userCollection.addUser(user);
        context.load(user);
        context.load(spGroup);

        context.executeQueryAsync(function() {
            INTRANET.Log(INTRANET.LogType.Info, "-----Inserted current user in admin group------");

        }, function(sender, args) {
            INTRANET.Log(INTRANET.LogType.Error, "Error in InsertCurrentUserSPGroup execute Query. Error: " + args);
        });
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in InsertCurrentUserSPGroup. Error: " + ex.message);
    }
};

//Set list as hidden
INTRANET.SetListToHidden = function() {
    try {
        INTRANET.Log(INTRANET.LogType.Info, "-----Setting list hidden------");
        var listName = INTRANET.App.ListBravoAppState;
        var hostWeb = hostWebContext.get_web();
        var list = hostWeb.get_lists().getByTitle(listName);
        list.set_hidden(true);
        list.Hidden = true;
        context.load(list);
        context.executeQueryAsync(
            function(result) {
                INTRANET.Log(INTRANET.LogType.Info, "-----List set to hidden------");
            },
            function onCreateListFailed(sender, args) {
                INTRANET.Log(INTRANET.LogType.Info, "SetListToHidden. Error: " + args.get_message());
            });
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in SetListToHidden. Error: " + ex.message);
    }
};

//Set list as hidden
INTRANET.GetAppStateItems = function() {
    try {
        var listName = INTRANET.App.ListBravoAppState;
        var hostWeb = hostWebContext.get_web();
        var list = hostWeb.get_lists().getByTitle(listName);
        var camlQuery = new SP.CamlQuery();
        camlQuery.set_viewXml('<View></View>');
        collAppStateItems = list.getItems(camlQuery);
        context.load(collAppStateItems);
        context.executeQueryAsync(
            function(result) {
                var listItemEnumerator = collAppStateItems.getEnumerator();
                var isInitialized = false;
                var isConfigured = false;

                while (listItemEnumerator.moveNext()) {
                    var oListItem = listItemEnumerator.get_current();
                    var itemKey = oListItem.get_item('BravoKey');
                    var itemValue = oListItem.get_item('BravoValue');

                    if (itemKey == INTRANET.App.APPState.Initialization && itemValue == INTRANET.App.Done) {
                        isInitialized = true;
                    } else if (itemKey == INTRANET.App.APPState.SiteConfiguration && itemValue == INTRANET.App.Done) {
                        isConfigured = true;
                    }
                }

                if (isConfigured == true) {
                    $(".configureInputs").hide();
                    $("#divRollBackInfo").show();
                    $("#configureRead").show();
                } else if (isInitialized == true) {
                    $(".configureInputs").hide();
                    $("#divRollBackInfo").show();
                    $("#configureRead").show();
                }
            },
            function onCreateListFailed(sender, args) {
                INTRANET.Log(INTRANET.LogType.Info, "GetAppStateItems. Error: " + args.get_message());
            });
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in GetAppStateItems. Error: " + ex.message);
    }
};

//Logging
INTRANET.Log = function(pLogType, PMessage) {
    try {
        if (pLogType == INTRANET.LogType.Error) {
            console.error(PMessage);
        } else {
            console.log(PMessage);
        }

        $('#txtAreaProgress').append("\n" + PMessage);
        document.getElementById("txtAreaProgress").scrollTop = document.getElementById("txtAreaProgress").scrollHeight;
    } catch (ex) {
        console.log("Error in Log method: " + ex.message);
    }
};

//Set Local Storage
INTRANET.SetHostWebURlInSession();

INTRANET.GetFilenameFromUrl = function(url) {
    var filename = url.substring(url.lastIndexOf('/') + 1);
    return filename;
};

INTRANET.GetPathFromUrl = function(url) {
    var path = url.substring(1, url.lastIndexOf('/') + 1);
    return path;
};

INTRANET.ArrayBufferToBase64 = function(buffer) {
    var binary = '';
    var bytes = new Uint8Array(buffer);
    var len = bytes.byteLength;
    for (var i = 0; i < len; i++) {
        binary += String.fromCharCode(bytes[i]);
    }
    return window.btoa(binary);
}

INTRANET.GetRelativeUrlFromAbsolute = function(absoluteUrl) {
    absoluteUrl = absoluteUrl.replace('https://', '');
    var parts = absoluteUrl.split('/');
    var relativeUrl = '/';

    for (var i = 1; i < parts.length; i++) {
        relativeUrl += parts[i] + '/';
    }
    return relativeUrl;
};

// Extends jquery ajaxTransport too support binary reading of files
$.ajaxTransport("+binary", function(options, originalOptions, jqXHR) {
    // check for conditions and support for blob / arraybuffer response type
    if (window.FormData && ((options.dataType && (options.dataType == 'binary')) || (options.data && ((window.ArrayBuffer && options.data instanceof ArrayBuffer) || (window.Blob && options.data instanceof Blob))))) {
        return {
            // create new XMLHttpRequest
            send: function(headers, callback) {
                // setup all variables
                var xhr = new XMLHttpRequest(),
                    url = options.url,
                    type = options.type,
                    async = options.async || true,
                    // blob or arraybuffer. Default is blob
                    dataType = options.responseType || "blob",
                    data = options.data || null,
                    username = options.username || null,
                    password = options.password || null;

                xhr.addEventListener('load', function() {
                    var data = {};
                    data[options.dataType] = xhr.response;
                    // make callback and send data
                    callback(xhr.status, xhr.statusText, data, xhr.getAllResponseHeaders());
                });

                xhr.open(type, url, async, username, password);

                // setup custom headers
                for (var i in headers) {
                    xhr.setRequestHeader(i, headers[i]);
                }

                xhr.responseType = dataType;
                xhr.send(data);
            },
            abort: function() {
                jqXHR.abort();
            }
        };
    }
});

/////////////////////////////////////////////Rollback///////////////////////////////////////

INTRANET.Rollback = function() {
    $("#configureRead").hide();
    $("#idInstallRollbackText").show();
    $("#idInstallRollbackCount").show();
    $("#txtAreaProgress").css("display", "block");
    $("#rollBackBtn").prop('disabled', true);
    $("#rollBackBtn").css('opacity', "0.65");

    INTRANET.Log(INTRANET.LogType.Info, "---RollBack Started---");
    INTRANET.IncreaseProgressCounter("RollBack Started");

    INTRANET.App.IsRollback = true;
    counterInstallation = 0;
    counterDeletedLists = 0;
    INTRANET.App.CollLists = [];
    collRemoveContentType = [];
    collRemoveFields = [];
    INTRANET.InIt();
    INTRANET.DeleteLists();
};

//Delete Custom Lists
INTRANET.DeleteLists = function() {
    try {
        INTRANET.Log(INTRANET.LogType.Info, "-----Delete Lists------");
        INTRANET.IncreaseProgressCounter("Removing List...");

        if (typeof(hostWebContext) != 'undefined' && hostWebContext != null) {

            INTRANET.Log(INTRANET.LogType.Info, "List deleting: " + INTRANET.Schema.Lists[counterDeletedLists].Name);

            var hostWeb = hostWebContext.get_web();
            var list = hostWeb.get_lists().getByTitle(INTRANET.Schema.Lists[counterDeletedLists].Name);
            list.deleteObject();

            context.load(list);
            context.executeQueryAsync(
                function() {
                    counterDeletedLists++;

                    if (INTRANET.Schema.Lists.length > counterDeletedLists) {
                        INTRANET.DeleteLists();
                    } else {
                        INTRANET.FetchContentTypes();
                    }
                },
                function onCreateListFailed(sender, args) {
                    counterDeletedLists++;

                    if (INTRANET.Schema.Lists.length > counterDeletedLists) {
                        INTRANET.DeleteLists();
                    } else {
                        INTRANET.FetchContentTypes();
                    }
                });
        }
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in DeleteLists. Error: " + ex.message);
    }
};

//Delete Content Types to List
INTRANET.FetchContentTypes = function() {
    try {
        counterDeletedLists = 0;
        INTRANET.Log(INTRANET.LogType.Info, "-----Fetching content type to Lists------");
        INTRANET.IncreaseProgressCounter("Removing Content types...");

        var hostWeb = hostWebContext.get_web();
        collContentTypes = hostWeb.get_contentTypes();
        //var contentType = contentTypeCol.getById(objList.ContentTypeID);

        context.load(collContentTypes);
        context.executeQueryAsync(
            function() {
                debugger
                var ctypesInfo = '';
                var ctypesEnumerator = collContentTypes.getEnumerator();

                while (ctypesEnumerator.moveNext()) {
                    var ocontentType = ctypesEnumerator.get_current();

                    if (ocontentType.get_group() == INTRANET.Constant.ContentTypeGroupName) {
                        var obj = {
                            Name: ocontentType.get_name(),
                            ID: ocontentType.get_id()
                        };
                        collRemoveContentType.push(obj);
                    }
                }
                INTRANET.Log(INTRANET.LogType.Info, "---Content Type Fetched---");

                INTRANET.DeleteContentType();
            },
            function(sender, args) {
                INTRANET.Log(INTRANET.LogType.Error, "FetchContentTypes. Error: " + args.get_message());
            });
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in FetchContentTypes. Error: " + ex.message);
    }
}

INTRANET.DeleteContentType = function() {
    try {

        if (collRemoveContentType.length > 0) {
            INTRANET.Log(INTRANET.LogType.Info, "Deleting Content Type: " + collRemoveContentType[counterDeletedLists].Name);

            var hostWeb = hostWebContext.get_web();
            var contentTypes = hostWeb.get_contentTypes();
            var contentType = contentTypes.getById(collRemoveContentType[counterDeletedLists].ID);
            contentType.deleteObject();

            context.load(contentType);
            context.executeQueryAsync(
                function() {
                    debugger
                    counterDeletedLists++;
                    if (collRemoveContentType.length > counterDeletedLists) {
                        INTRANET.DeleteContentType();
                    } else {
                        INTRANET.FetchFields();
                    }
                },
                function(sender, args) {
                    counterDeletedLists++;
                    if (collRemoveContentType.length > counterDeletedLists) {
                        INTRANET.DeleteContentType();
                    } else {
                        INTRANET.FetchFields();
                    }
                });
        } else {
            INTRANET.FetchFields();
        }
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in DeleteContentType. Error: " + ex.message);
    }
};

// Fetch fields 
INTRANET.FetchFields = function() {
    try {
        counterDeletedLists = 0;
        INTRANET.Log(INTRANET.LogType.Info, "-----Fetching Site Columns------");
        INTRANET.IncreaseProgressCounter("Removing fields...");

        var hostWeb = hostWebContext.get_web();
        collFields = hostWeb.get_fields();

        context.load(collFields);
        context.executeQueryAsync(
            function() {
                debugger
                var fieldInfo = '';
                var fieldsEnumerator = collFields.getEnumerator();

                while (fieldsEnumerator.moveNext()) {
                    var field = fieldsEnumerator.get_current();

                    if (field.get_group() == INTRANET.Constant.SiteColumnGroupName) {
                        var obj = {
                            Name: field.get_title(),
                            ID: field.get_id()
                        };
                        collRemoveFields.push(obj);
                    }
                }
                INTRANET.Log(INTRANET.LogType.Info, "-----Fetched Site Columns------");

                INTRANET.DeleteField();
            },
            function(sender, args) {
                INTRANET.Log(INTRANET.LogType.Error, "FetchFields. Error: " + args.get_message());
            });
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in FetchFields. Error: " + ex.message);
    }
};

INTRANET.DeleteField = function() {
    try {
        if (collRemoveFields.length > 0) {
            INTRANET.Log(INTRANET.LogType.Info, "Removing Site Column: " + collRemoveFields[counterDeletedLists].Name);

            var hostWeb = hostWebContext.get_web();
            var fields = hostWeb.get_fields();
            var field = fields.getById(collRemoveFields[counterDeletedLists].ID);
            field.deleteObject();

            context.load(field);
            context.executeQueryAsync(
                function() {
                    debugger
                    counterDeletedLists++;
                    if (collRemoveFields.length > counterDeletedLists) {
                        INTRANET.DeleteField();
                    } else {
                        INTRANET.EndRollBack();
                    }
                },
                function(sender, args) {
                    counterDeletedLists++;
                    if (collRemoveFields.length > counterDeletedLists) {
                        INTRANET.DeleteField();
                    } else {
                        INTRANET.EndRollBack();
                    }
                });
        } else {
            INTRANET.EndRollBack();
        }
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in DeleteField. Error: " + ex.message);
    }
};

INTRANET.DeleteSPGroup = function() {
    try {
        var hostWeb = hostWebContext.get_web();
        var collGroup = hostWeb.get_siteGroups();
        var oGroup = collGroup.getByName(INTRANET.App.SPGroup.Name);
        collGroup.remove(oGroup);

        context.executeQueryAsync(
            function() {
                debugger
                INTRANET.EndRollBack();
            },
            function(sender, args) {
                INTRANET.EndRollBack();
            });
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in DeleteSPGroup. Error: " + ex.message);
    }
};

INTRANET.EndRollBack = function() {
    INTRANET.Log(INTRANET.LogType.Info, "---RollBack Completed---");
    INTRANET.IncreaseProgressCounter("RollBack Completed!");

    $("#idInstallRollbackText").hide();
    $("#idInstallRollbackCount").hide();

    $("#idRollBackConfirmation").show();
};

var licenseCollection;
var licenseResponse;
//Retrieve License
INTRANET.RetrieveLicense = function(productID) {
    try {
        INTRANET.Log(INTRANET.LogType.Info, "---Retrieving Licensing Details---");

        var context = SP.ClientContext.get_current();

        //Retrieve license from SharePoint; change this productId with the one for your app; you can get it from AppManifest.xml
        licenseCollection = SP.Utilities.Utility.getAppLicenseInformation(context, productID);
        context.executeQueryAsync(
            function() {
                INTRANET.Log(INTRANET.LogType.Info, "---License Token fetched: RetrieveLicense---");

                var topLicense = licenseCollection.get_item(0).get_rawXMLLicenseToken();
                //encode license; required to call the verification service since it will be sent on the URL 
                var encodedTopLicense = encodeURIComponent(topLicense);
                INTRANET.VerifyLicense(encodedTopLicense);

            },
            function(sender, args) {
                var errMessage = 'Request failed in RetrieveLicense. ' + args.get_message() + '\n' + args.get_stackTrace();
                INTRANET.Log(INTRANET.LogType.Error, "Error in RetrieveLicense. Error: " + errMessage);
            });
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in RetrieveLicense. Error: " + ex.message);
    }
};

//Verify user license
INTRANET.VerifyLicense = function(encodedTopLicense) {

    try {
        var context = SP.ClientContext.get_current();
        //Call verification service via the WebProxy; this ensures the validity of the license in case SharePoint was tampered with
        var request = new SP.WebRequestInfo();
        //We use the REST flavor of the verification service
        request.set_url(INTRANET.App.LicenseVerificationService + encodedTopLicense);
        request.set_method("GET");
        licenseResponse = SP.WebProxy.invoke(context, request);
        // Set the event handlers and invoke the request
        context.executeQueryAsync(
            function() {
                INTRANET.Log(INTRANET.LogType.Info, "---License Token fetched: VerifyLicense---");

                try {
                    //Start - License Response                            
                    var verificationResponse = licenseResponse.get_body();
                    var xmlDoc = $.parseXML(verificationResponse);
                    var $xml = $(xmlDoc);
                    var licenseType = $xml.find("EntitlementType").text();
                    var licenseIsValid = $xml.find("IsValid").text();
                    var licenseIsTest = $xml.find("IsTest").text();

                    INTRANET.Constant.LicenseType = licenseType;

                    //DO SOMETHING NOW THAT YOU KNOW WHAT TYPE OF LICENSE THE USER HAS
                    switch (licenseType) {
                        case "Free":
                            //alert("Free app");
                            break;
                        case "Paid":
                            //alert("Paid app");
                            break;
                        case "Trial":
                            INTRANET.Constant.LicenseExpiryDate = $xml.find("EntitlementExpiryDate").text();
                            //alert("Trial app");
                            //You can then look at the expiration date on the response
                            break;
                    }
                } catch (err) {
                    INTRANET.Log(INTRANET.LogType.Error, "Error in VerifyLicense Fetch details. Error: " + ex.message);
                }
                //Create Tenant details
                INTRANET.CreateTenantInSQL();

            },
            function(sender, args) {
                var errMessage = 'Request failed in VerifyLicense. ' + args.get_message() + '\n' + args.get_stackTrace();
                INTRANET.Log(INTRANET.LogType.Error, "Error in VerifyLicense. Error: " + errMessage);
            });
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in VerifyLicense. Error: " + ex.message);
    }
}

INTRANET.UploadMasterPage = function() {
    INTRANET.ReadFromAppWebAndProvisionToHost(_spPageContextInfo.webAbsoluteUrl +
        '/IntranetMasterPagesFiles/MasterPage.txt',
        '_catalogs/masterpage', 'NBB_Site.master').then(function() {
        INTRANET.SetAsDefaultMasterPage(_spPageContextInfo.siteServerRelativeUrl +
            '/_catalogs/masterpage/NBB_Site.master');

    });
}

INTRANET.ReadFromAppWebAndProvisionToHost = function(appPageUrl, folderPath, hostWebFileName) {
    var deferred = $.Deferred();
    var destinationServerRelativeUrl = hostWebUrl,
        destinationFileName = hostWebFileName;
    var req = $.ajax({
        url: appPageUrl,
        type: "GET",
        cache: false
    }).done(function(fileContents) {
        if (fileContents !== undefined && fileContents.length > 0) {
            INTRANET.UploadFileToHostWebViaCSOM(folderPath, destinationFileName, fileContents).then(function() {
                deferred.resolve();
            }, function() {
                deferred.reject();
            });
        } else {
            alert('Failed to read file from app web, so not uploading to host web..');
        }

        deferred.resolve();
    }).fail(function(jqXHR, textStatus) {
        alert("Request for page in app web failed: " + textStatus);
        deferred.reject();
    });

    return deferred.promise();
}

INTRANET.UploadFileToHostWebViaCSOM = function(folderPath, filename, contents) {
    var deferred = $.Deferred();
    var createInfo = new SP.FileCreationInformation();
    createInfo.set_content(new SP.Base64EncodedByteArray());

    for (var i = 0; i < contents.length; i++) {

        createInfo.get_content().append(contents.charCodeAt(i));
    }
    createInfo.set_overwrite(true);
    createInfo.set_url(filename);
    var hostWeb = hostWebContext.get_web();
    var files = hostWeb.getFolderByServerRelativeUrl(INTRANET.GetRelativeUrlFromAbsolute(hostWebUrl) + folderPath).get_files();
    context.load(files);

    files.add(createInfo);
    context.executeQueryAsync(function(data) {
            INTRANET.PublishFile(INTRANET.GetRelativeUrlFromAbsolute(hostWebUrl) + folderPath + '/' + filename).then(function() {
                console.log('File uploaded successfully: ' + INTRANET.GetRelativeUrlFromAbsolute(hostWebUrl) + '/' + filename);
                deferred.resolve();
            }, function(error) {
                deferred.reject();
            });
        },
        function(sender, args) {
            console.log('Failed uploadation fail:' + sender.statusCode);
        });

    return deferred.promise();
}

INTRANET.PublishFile = function(fileRelativeUrl) {
    var deferred = $.Deferred();
    var file = hostWeb.getFileByServerRelativeUrl(fileRelativeUrl);
    file.publish();
    context.load(file);
    context.executeQueryAsync(function(data) {
        deferred.resolve();
    }, function(err, sender) {
        console.log('Error publishing file:', err, sender);
    });
    return deferred.promise();
}

INTRANET.SetAsDefaultMasterPage = function(masterPageUrl) {
    var hostWeb = hostWebContext.get_web();
    hostWeb.set_customMasterUrl(masterPageUrl);
    //hostWeb.set_masterUrl(masterPageUrl);
    hostWeb.update();

    context.load(hostWeb);
    context.executeQueryAsync(function(data) {
            console.log('Master page updated successfully..');
        },
        function(sender, args) {
            alert('Failed to update master page on host web. Error:' + args.get_message());
        });
}

// INTRANET.SetAsDefaultMasterPage("/sites/IntranetInBoxPublishing/_catalogs/masterpage/NBB_Site.master")

INTRANET.UploadAllJSFiles = function() {
    try {

        var CSSFolderPath = INTRANET.GetRelativeUrlFromAbsolute(hostWebUrl) + "Style Library/NBB",
            tempFileArray = INTRANET.Schema.AllJS,
            path = INTRANET.GetRelativeUrlFromAbsolute(hostWebUrl) + "Style Library/NBB/JS/";
        INTRANET.Log(INTRANET.LogType.Info, "-----Uploading JS Files-----");

        var asyncmethods = [];
        if (tempFileArray) {
            asyncmethods = $.map(tempFileArray, function(fileName) {
                var txtName = fileName + ".js",
                    fileNameUrl = _spPageContextInfo.webAbsoluteUrl + "/IntranetJSFiles/" + fileName + ".txt",
                    readFile = $.ajax({
                        url: fileNameUrl,
                        type: "GET",
                        cache: false,
                        success: function(fileContents) {
                            if (fileContents !== undefined && fileContents.length > 0) {
                                INTRANET.UploadFile(txtName, fileContents, path);
                            }
                        },
                        error: function(error) {
                            console.log(error);
                        }
                    });
            });
        }
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in Uploading JS files. Error: " + ex.message);
    }
}

INTRANET.UploadAllCSSFiles = function() {
    try {

        var CSSFolderPath = INTRANET.GetRelativeUrlFromAbsolute(hostWebUrl) + "Style Library/NBB",
            tempCSSArray = INTRANET.Schema.AllCSS,
            path = INTRANET.GetRelativeUrlFromAbsolute(hostWebUrl) + "Style Library/NBB/CSS/";
        INTRANET.Log(INTRANET.LogType.Info, "-----Uploading CSS Files-----");

        var asyncmethods = [];
        if (tempCSSArray) {
            asyncmethods = $.map(tempCSSArray, function(cssName) {
                var txtName = cssName + ".css",
                    cssNameUrl = _spPageContextInfo.webAbsoluteUrl + "/IntranetCSS/" + cssName + ".txt",
                    readFile = $.ajax({
                        url: cssNameUrl,
                        type: "GET",
                        cache: false,
                        success: function(fileContents) {
                            if (fileContents !== undefined && fileContents.length > 0) {
                                INTRANET.UploadFile(txtName, fileContents, path);
                            }
                        },
                        error: function(error) {
                            console.log(error);
                        }
                    });
            });
        }


    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in Uploading css files. Error: " + ex.message);
    }

}

INTRANET.UploadAllTextFiles = function() {
    try {

        var tempFileArray = INTRANET.Schema.AllTextFiles,
            path = INTRANET.GetRelativeUrlFromAbsolute(hostWebUrl) + "Style Library/NBB/Text/";
        INTRANET.Log(INTRANET.LogType.Info, "-----Uploading Text Files-----");

        var asyncmethods = [];
        if (tempFileArray) {
            asyncmethods = $.map(tempFileArray, function(fileName) {
                var txtName = fileName,
                    fileNameUrl = _spPageContextInfo.webAbsoluteUrl + "/IntranetTextFiles/" + fileName,
                    readFile = $.ajax({
                        url: fileNameUrl,
                        type: "GET",
                        cache: false,
                        success: function(fileContents) {
                            if (fileContents !== undefined && fileContents.length > 0) {
                                INTRANET.UploadFile(txtName, fileContents, path);
                            }
                        },
                        error: function(error) {
                            console.log(error);
                        }
                    });
            });
        }
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in Uploading Text files. Error: " + ex.message);
    }
}

INTRANET.UploadAllPageFiles = function(listtitle, fileArray, uploadUrl, type, readUrl) {
    var deferred = $.Deferred();
    try {
        var tempFileArray = fileArray,
            url = uploadUrl,
            hostWebUrl = hostWeb.get_url(),
            hostWebRelativeUrl = hostWeb.get_serverRelativeUrl();
        INTRANET.Log(INTRANET.LogType.Info, type + " Files Uploading");
        var asyncmethods = [];
        if (tempFileArray) {
            asyncmethods = $.map(tempFileArray, function(fileName) {
                var txtName = fileName + "." + type,
                    fileNameUrl = readUrl + fileName + ".txt",
                    dfd = $.Deferred();
                $.ajax({
                    url: fileNameUrl,
                    type: "GET",
                    cache: false,
                    success: function(fileContents) {
                        if (fileContents !== undefined && fileContents.length > 0) {
                            fileContents = fileContents.replace("{site_absolute_url}", hostWebUrl)
                                .replace("{site_relative_url}", hostWebRelativeUrl);
                            INTRANET.UploadFile(listtitle, txtName, fileContents, url).then(function() {
                                dfd.resolve();
                            }, function(err) {
                                dfd.reject(err);
                            });
                        }
                    },
                    error: function(error) {
                        console.log(error);
                        dfd.reject(err);
                    }
                });
                return dfd.promise();
            });

            $.when.apply($, asyncmethods).done(function(results) {
                deferred.resolve();
            });
        }
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in Uploading files. Error: " + ex.message);
    }

    return deferred.promise();
}

INTRANET.UploadAllFiles = function(listTitle, fileArray, uploadUrl, readUrl) {
    var deferred = $.Deferred();
    try {
        var tempFileArray = fileArray,
            url = uploadUrl;
        INTRANET.Log(INTRANET.LogType.Info + " Files Uploading");
        var asyncmethods = [];
        if (tempFileArray) {
            asyncmethods = $.map(tempFileArray, function(fileName) {
                var txtName = fileName,
                    fileNameUrl = readUrl + fileName,
                    readFile = $.ajax({
                        url: fileNameUrl,
                        type: "GET",
                        cache: false,
                        dataType: "binary",
                        processData: false,
                        responseType: 'arraybuffer',
                        success: function(fileContents) {
                            INTRANET.UploadFile(listTitle, txtName, fileContents, url, true);
                        },
                        error: function(error) {
                            console.log(error);
                        }
                    });
            });
        }

        $.when.apply($, asyncMethods).done(function(results) {
            deferred.resolve();
        }, function(err) {
            deferred.reject();
        });
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in Uploading Text files. Error: " + ex.message);
    }

    return deferred.promise();
}

INTRANET.CreateAllFolders = function(listTitle, rootFolderName) {
    var deferred = $.Deferred();
    try {
        var list = hostWeb.get_lists().getByTitle(listTitle);
        INTRANET.Log(INTRANET.LogType.Info, "-----creating Folders---");
        var createFolderInternal = function(parentFolder, folderName) {
            var dfd = $.Deferred();
            var ctx = parentFolder.get_context();
            var curFolder = parentFolder.get_folders().add(folderName);
            ctx.load(curFolder);
            ctx.executeQueryAsync(
                function() {
                    INTRANET.Log(INTRANET.LogType.Info, "Folder Created." + curFolder);
                    dfd.resolve(curFolder);
                },
                function(sender, args) {
                    INTRANET.Log(INTRANET.LogType.Info, "Error in creatin a folder. Error:" + args.get_message);
                    dfd.reject();
                });
            return dfd.promise();
        };
        createFolderInternal(list.get_rootFolder(), rootFolderName).then(function(curFolder) {
            $.when(createFolderInternal(curFolder, "css"),
                createFolderInternal(curFolder, "js"),
                createFolderInternal(curFolder, "html"),
                createFolderInternal(curFolder, "text"),
                createFolderInternal(curFolder, "fonts"),
                createFolderInternal(curFolder, "images")).done(function() {
                deferred.resolve();
            });
        });
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in creating folder. Error: " + ex.message);
        deferred.reject();
    }

    return deferred.promise();

};

INTRANET.UploadFile = function(listtitle, filename, contents, path, isArrayBuffer) {
    var deferred = $.Deferred();
    try {
        var createInfo = new SP.FileCreationInformation();

        createInfo.set_url(path + filename);
        createInfo.set_overwrite(true);

        var list = hostWeb.get_lists().getByTitle(listtitle);

        if (!isArrayBuffer) {
            createInfo.set_content(new SP.Base64EncodedByteArray());
            for (var i = 0; i < contents.length; i++) {

                createInfo.get_content().append(contents.charCodeAt(i));
            }
        } else {
            createInfo.set_content(INTRANET.ArrayBufferToBase64(contents));
        }


        this.newFile = list.get_rootFolder().get_files().add(createInfo);
        this.newFile.checkIn();
        this.newFile.publish();
        context.load(this.newFile);
        context.executeQueryAsync(function(data) {
                INTRANET.Log(INTRANET.LogType.Info, 'File uploaded successfully: ' + path + filename);
                deferred.resolve();
            },
            function(sender, args) {
                INTRANET.Log(INTRANET.LogType.Info, 'File upload Failed. Error: ' + args.get_message);
                deferred.reject();
            });
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Info, 'Error in uploading File. Error: ' + ex.message);
    }

    return deferred.promise();
}


INTRANET.UploadAllDisplayTemplates = function(fileArray, uploadUrl, type, readUrl) {
    var deferred = $.Deferred();
    try {
        var tempFileArray = fileArray,
            url = uploadUrl;
        INTRANET.Log(INTRANET.LogType.Info, type + " Files Uploading");
        var asyncmethods = [];
        if (tempFileArray) {
            asyncmethods = $.map(tempFileArray, function(fileName) {
                var dfd = $.Deferred(),
                    txtName = fileName + "." + type,
                    fileNameUrl = readUrl + fileName + ".txt",
                    readFile = $.ajax({
                        url: fileNameUrl,
                        type: "GET",
                        cache: false,
                        success: function(fileContents) {
                            if (fileContents !== undefined && fileContents.length > 0) {
                                INTRANET.UploadFileToHostWebViaCSOM(url, txtName, fileContents).then(function() {
                                    dfd.resolve();
                                });
                            }
                        },
                        error: function(error) {
                            console.log(error);
                        }
                    });

                return dfd.promise();
            });

            $.when.apply($, asyncmethods).done(function(results) {
                deferred.resolve();
            });

        }
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in Uploading Text files. Error: " + ex.message);
    }

    return deferred.promise();
}


INTRANET.UploadAllPageLayouts = function(fileArray, uploadUrl, type, readUrl) {
    var deferred = $.Deferred();
    try {
        var tempFileArray = fileArray,
            url = uploadUrl;
        INTRANET.Log(INTRANET.LogType.Info, type + " Files Uploading");
        var asyncmethods = [];
        if (tempFileArray) {
            asyncmethods = $.map(tempFileArray, function(fileName) {
                var txtName = fileName + "." + type,
                    fileNameUrl = readUrl + fileName + ".txt",
                    dfd = $.Deferred();
                $.ajax({
                    url: fileNameUrl,
                    type: "GET",
                    cache: false,
                    success: function(fileContents) {
                        if (fileContents !== undefined && fileContents.length > 0) {
                            INTRANET.UploadFileToHostWebViaCSOM(url, txtName, fileContents).then(function() {
                                INTRANET.SetPageLayoutContentType(INTRANET.GetRelativeUrlFromAbsolute(hostWebUrl) + url + '/' + txtName, txtName).then(function() {
                                    dfd.resolve();
                                }, function(err) {
                                    console.log(error);
                                    dfd.resolve();
                                });
                            });
                        }
                    },
                    error: function(error) {
                        console.log(error);
                        dfd.resolve();
                    }
                });

                return dfd.promise();
            });

            $.when.apply($, asyncmethods).done(function(results) {
                deferred.resolve();
            });
        }
    } catch (ex) {
        INTRANET.Log(INTRANET.LogType.Error, "Error in Uploading Text files. Error: " + ex.message);
    }
    return deferred.promise();
}


INTRANET.SetPageLayoutContentType = function(pageLayoutUrl, dtName) {
    var deferred = $.Deferred();
    var file = hostWeb.getFileByServerRelativeUrl(pageLayoutUrl);
    var item = file.get_listItemAllFields();;
    context.load(item);
    context.executeQueryAsync(function(data) {
        item["ContentTypeId"] = "0x01010007FF3E057FA8AB4AA42FCB67B453FFC100E214EEE741181F4E9F7ACC43278EE8110003D357F861E29844953D5CAA1D4D8A3B",
            item["PublishingAssociatedContentType"] = "#Article Page;#0x010100C568DB52D9D0A14D9B2FDCC96666E9F2007948130EC3DB064584E219954237AF3900242457EFB8B24247815D688C526CD44D;#",
            item.update();
        context.executeQueryAsync(function(data) {
            deferred.resolve();
        }, function(err, sender) {
            deferred.reject(err.sender);
        });
    }, function(err, sender) {
        deferred.resolve(err, sender);
    });
    return deferred.resolve();
};
INTRANET.AddWebPartsToPage = function(serverRelativeUrl) {
    var oFile = hostWebContext.get_web().getFileByServerRelativeUrl(serverRelativeUrl);
    oFile.checkOut();
    var limitedWebPartManager = oFile.getLimitedWebPartManager(SP.WebParts.PersonalizationScope.shared);
    var orgAnnouncementWebPartXml = `<?xml version="1.0" encoding="utf-8"?>
    <WebPart xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://schemas.microsoft.com/WebPart/v2">
      <Title>Org Announcement carousel CEWP</Title>
      <FrameType>None</FrameType>
      <Description>Allows authors to enter rich text content.</Description>
      <IsIncluded>true</IsIncluded>
      <ZoneID>carousel</ZoneID>
      <PartOrder>4</PartOrder>
      <FrameState>Normal</FrameState>
      <Height />
      <Width />
      <AllowRemove>true</AllowRemove>
      <AllowZoneChange>true</AllowZoneChange>
      <AllowMinimize>true</AllowMinimize>
      <AllowConnect>true</AllowConnect>
      <AllowEdit>true</AllowEdit>
      <AllowHide>true</AllowHide>
      <IsVisible>true</IsVisible>
      <DetailLink />
      <HelpLink />
      <HelpMode>Modeless</HelpMode>
      <Dir>Default</Dir>
      <PartImageSmall />
      <MissingAssembly>Cannot import this Web Part.</MissingAssembly>
      <PartImageLarge>/sites/IntranetInBoxPublishing/_layouts/15/images/mscontl.gif</PartImageLarge>
      <IsIncludedFilter />
      <IsAvailable>true</IsAvailable>
      <Assembly>Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c</Assembly>
      <TypeName>Microsoft.SharePoint.WebPartPages.ContentEditorWebPart</TypeName>
      <ContentLink xmlns="http://schemas.microsoft.com/WebPart/v2/ContentEditor">{site_relative_url}/Style Library/IntranetInbox/html/carousel.html</ContentLink>
      <Content xmlns="http://schemas.microsoft.com/WebPart/v2/ContentEditor" />
      <PartStorage xmlns="http://schemas.microsoft.com/WebPart/v2/ContentEditor" />
    </WebPart>`;

    var pollWebPartXml = `<?xml version="1.0" encoding="utf-8"?>
    <WebPart xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://schemas.microsoft.com/WebPart/v2">
      <Title>Poll content editor</Title>
      <FrameType>None</FrameType>
      <Description>Allows authors to enter rich text content.</Description>
      <IsIncluded>true</IsIncluded>
      <ZoneID>poll</ZoneID>
      <PartOrder>0</PartOrder>
      <FrameState>Normal</FrameState>
      <Height />
      <Width />
      <AllowRemove>true</AllowRemove>
      <AllowZoneChange>true</AllowZoneChange>
      <AllowMinimize>true</AllowMinimize>
      <AllowConnect>true</AllowConnect>
      <AllowEdit>true</AllowEdit>
      <AllowHide>true</AllowHide>
      <IsVisible>true</IsVisible>
      <DetailLink />
      <HelpLink />
      <HelpMode>Modeless</HelpMode>
      <Dir>Default</Dir>
      <PartImageSmall />
      <MissingAssembly>Cannot import this Web Part.</MissingAssembly>
      <PartImageLarge>/sites/IntranetInBoxPublishing/_layouts/15/images/mscontl.gif</PartImageLarge>
      <IsIncludedFilter />
      <IsAvailable>true</IsAvailable>
      <Assembly>Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c</Assembly>
      <TypeName>Microsoft.SharePoint.WebPartPages.ContentEditorWebPart</TypeName>
      <ContentLink xmlns="http://schemas.microsoft.com/WebPart/v2/ContentEditor">{site_relative_url}/Style Library/IntranetInbox/html/Poll.html</ContentLink>
      <Content xmlns="http://schemas.microsoft.com/WebPart/v2/ContentEditor" />
      <PartStorage xmlns="http://schemas.microsoft.com/WebPart/v2/ContentEditor" />
    </WebPart>`;

    var birthDayWebPartXml = `<?xml version="1.0" encoding="utf-8"?>
    <WebPart xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://schemas.microsoft.com/WebPart/v2">
      <Title>Birthday And Anniversaries</Title>
      <FrameType>None</FrameType>
      <Description>Allows authors to enter rich text content.</Description>
      <IsIncluded>true</IsIncluded>
      <ZoneID>x450228fd78454faaa9c3cb246991d0f7</ZoneID>
      <PartOrder>0</PartOrder>
      <FrameState>Normal</FrameState>
      <Height />
      <Width />
      <AllowRemove>true</AllowRemove>
      <AllowZoneChange>true</AllowZoneChange>
      <AllowMinimize>true</AllowMinimize>
      <AllowConnect>true</AllowConnect>
      <AllowEdit>true</AllowEdit>
      <AllowHide>true</AllowHide>
      <IsVisible>true</IsVisible>
      <DetailLink />
      <HelpLink />
      <HelpMode>Modeless</HelpMode>
      <Dir>Default</Dir>
      <PartImageSmall />
      <MissingAssembly>Cannot import this Web Part.</MissingAssembly>
      <PartImageLarge>/_layouts/15/images/mscontl.gif</PartImageLarge>
      <IsIncludedFilter />
      <Assembly>Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c</Assembly>
      <TypeName>Microsoft.SharePoint.WebPartPages.ContentEditorWebPart</TypeName>
      <ContentLink xmlns="http://schemas.microsoft.com/WebPart/v2/ContentEditor">{site_relative_url}/Style Library/IntranetInbox/html/BirthdayAndAnniversaries.html</ContentLink>
      <Content xmlns="http://schemas.microsoft.com/WebPart/v2/ContentEditor"><![CDATA[​​<br/><br/>]]></Content>
      <PartStorage xmlns="http://schemas.microsoft.com/WebPart/v2/ContentEditor" />
    </WebPart>`;

    var newsWebPartXml = `<webParts>
    <webPart xmlns="http://schemas.microsoft.com/WebPart/v3">
      <metaData>
        <type name="Microsoft.Office.Server.Search.WebControls.ContentBySearchWebPart, Microsoft.Office.Server.Search, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" />
        <importErrorMessage>Cannot import this Web Part.</importErrorMessage>
      </metaData>
      <data>
        <properties>
          <property name="BypassResultTypes" type="bool">True</property>
          <property name="ItemTemplateId" type="string">~sitecollection/_catalogs/masterpage/Display Templates/Content Web Parts/Item_News.js</property>
          <property name="PropertyMappings" type="string" />
          <property name="ChromeState" type="chromestate">Normal</property>
          <property name="IncludeResultTypeConstraint" type="bool">False</property>
          <property name="StartingItemIndex" type="int">1</property>
          <property name="ShowDefinitions" type="bool">False</property>
          <property name="Height" type="string" />
          <property name="Hidden" type="bool">False</property>
          <property name="HitHighlightedPropertiesJson" type="string">["Title","Path","Author","SectionNames","SiteDescription"]</property>
          <property name="ScrollToTopOnRedraw" type="bool">False</property>
          <property name="UseSharedDataProvider" type="bool">False</property>
          <property name="RepositionLanguageDropDown" type="bool">False</property>
          <property name="AlwaysRenderOnServer" type="bool">False</property>
          <property name="AllowConnect" type="bool">True</property>
          <property name="ItemBodyTemplateId" type="string" />
          <property name="ShowAlertMe" type="bool">True</property>
          <property name="ExportMode" type="exportmode">All</property>
          <property name="AddSEOPropertiesFromSearch" type="bool">False</property>
          <property name="ShowUpScopeMessage" type="bool">False</property>
          <property name="AllowHide" type="bool">True</property>
          <property name="AllowClose" type="bool">True</property>
          <property name="UseSimplifiedQueryBuilder" type="bool">False</property>
          <property name="ShouldHideControlWhenEmpty" type="bool">False</property>
          <property name="ResultType" type="string" />
          <property name="LogAnalyticsViewEvent" type="bool">False</property>
          <property name="MaxPagesAfterCurrent" type="int">1</property>
          <property name="TitleUrl" type="string" />
          <property name="EmptyMessage" type="string" />
          <property name="AdvancedSearchPageAddress" type="string">advanced.aspx</property>
          <property name="IsXGeo3SForwardingFlighted" type="bool">True</property>
          <property name="AllowMinimize" type="bool">True</property>
          <property name="ShowBestBets" type="bool">False</property>
          <property name="AllowEdit" type="bool">True</property>
          <property name="NumberOfItems" type="int">7</property>
          <property name="HelpUrl" type="string" />
          <property name="ShowPaging" type="bool">True</property>
          <property name="ShowViewDuplicates" type="bool">False</property>
          <property name="SelectedPropertiesJson" type="string">["Path","Title","LastModifiedTime","NewsTitleOWSTEXT","NewsPublishedDateOWSDATE","MPLocation","NewsBannerOWSIMGE","NewsThumbnailOWSIMGE","MPDepartment"]</property>
          <property name="TargetResultTable" type="string">RelevantResults</property>
          <property name="HelpMode" type="helpmode">Modeless</property>
          <property name="ShowXGeoOptions" type="bool">False</property>
          <property name="IsXGeoFlighted" type="bool">False</property>
          <property name="ShowPersonalFavorites" type="bool">False</property>
          <property name="EnableXGeo3SForwarding" type="bool">False</property>
          <property name="PreloadedItemTemplateIdsJson" type="string">null</property>
          <property name="Description" type="string">Content Search Web Part will allow you to show items that are results of a search query you specify. When you add it to the page, this Web Part will show recently modified items from the current site. You can change this setting to show items from another site or list by editing the Web Part and changing its search criteria.As new content is discovered by search, this Web Part will display an updated list of items each time the page is viewed.</property>
          <property name="ShowPreferencesLink" type="bool">True</property>
          <property name="QueryGroupName" type="string">75ba926c-ec23-4184-b3d9-da70db9fab95</property>
          <property name="ShowResultCount" type="bool">True</property>
          <property name="TitleIconImageUrl" type="string" />
          <property name="Direction" type="direction">NotSet</property>
          <property name="ResultsPerPage" type="int">7</property>
          <property name="AvailableSortsJson" type="string">null</property>
          <property name="ShowResults" type="bool">True</property>
          <property name="ServerIncludeScriptsJson" type="string">null</property>
          <property name="SearchCenterXGeoLocations" type="string" />
          <property name="DataProviderJSON" type="string">{"QueryGroupName":"75ba926c-ec23-4184-b3d9-da70db9fab95","QueryPropertiesTemplateUrl":"sitesearch://webroot","IgnoreQueryPropertiesTemplateUrl":false,"SourceID":"8413cd39-2156-4e00-b54d-11efd9abdb89","SourceName":"Local SharePoint Results","SourceLevel":"Ssa","CollapseSpecification":"","QueryTemplate":"((path:{\\Site.URL}/Pages/) AND (ContentType:KPCU_News) AND (IsDocument:\"True\" OR contentclass:\"STS_ListItem\"))","FallbackSort":[{"p":"LastModifiedTime","d":1}],"FallbackSortJson":"[{\"p\":\"LastModifiedTime\",\"d\":1}]","RankRules":null,"RankRulesJson":"null","AsynchronousResultRetrieval":false,"SendContentBeforeQuery":true,"BatchClientQuery":true,"FallbackLanguage":-1,"FallbackRankingModelID":"","EnableStemming":true,"EnablePhonetic":false,"EnableNicknames":false,"EnableInterleaving":false,"EnableQueryRules":true,"EnableOrderingHitHighlightedProperty":false,"HitHighlightedMultivaluePropertyLimit":-1,"IgnoreContextualScope":true,"ScopeResultsToCurrentSite":false,"TrimDuplicates":false,"Properties":{"TryCache":true,"Scope":"{Site.URL}","UpdateLinksForCatalogItems":true,"EnableStacking":true,"CrossGeoQuery":"false","ListId":"55e27111-388b-411d-87df-351f2b92600f","ListItemId":4},"PropertiesJson":"{\"TryCache\":true,\"Scope\":\"{Site.URL}\",\"UpdateLinksForCatalogItems\":true,\"EnableStacking\":true,\"CrossGeoQuery\":\"false\",\"ListId\":\"55e27111-388b-411d-87df-351f2b92600f\",\"ListItemId\":4}","ClientType":"ContentSearchRegular","ClientFunction":"","ClientFunctionDetails":"","UpdateAjaxNavigate":true,"SummaryLength":180,"DesiredSnippetLength":90,"PersonalizedQuery":false,"FallbackRefinementFilters":null,"IgnoreStaleServerQuery":false,"RenderTemplateId":"DefaultDataProvider","AlternateErrorMessage":null,"Title":""}</property>
          <property name="ShowAdvancedLink" type="bool">True</property>
          <property name="ShowDidYouMean" type="bool">False</property>
          <property name="AllowZoneChange" type="bool">True</property>
          <property name="ChromeType" type="chrometype">None</property>
          <property name="GroupTemplateId" type="string">~sitecollection/_catalogs/masterpage/Display Templates/Content Web Parts/Group_Content.js</property>
          <property name="MissingAssembly" type="string">Cannot import this Web Part.</property>
          <property name="OverwriteResultPath" type="bool">True</property>
          <property name="Width" type="string" />
          <property name="MaxPagesBeforeCurrent" type="int">4</property>
          <property name="XGeoTenantsInfo" type="string" />
          <property name="ShowLanguageOptions" type="bool">True</property>
          <property name="ResultTypeId" type="string" />
          <property name="AlternateErrorMessage" type="string" null="true" />
          <property name="Title" type="string">KPCU News</property>
          <property name="RenderTemplateId" type="string">~sitecollection/_catalogs/masterpage/Display Templates/Content Web Parts/Control_KPCUNews.js</property>
          <property name="EmitStyleReference" type="bool">True</property>
          <property name="StatesJson" type="string">{}</property>
          <property name="ShowSortOptions" type="bool">False</property>
          <property name="CatalogIconImageUrl" type="string" />
        </properties>
      </data>
    </webPart>
  </webParts>`;

    var eventsWebPartXml = `<webParts>
  <webPart xmlns="http://schemas.microsoft.com/WebPart/v3">
    <metaData>
      <type name="Microsoft.Office.Server.Search.WebControls.ContentBySearchWebPart, Microsoft.Office.Server.Search, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" />
      <importErrorMessage>Cannot import this Web Part.</importErrorMessage>
    </metaData>
    <data>
      <properties>
        <property name="BypassResultTypes" type="bool">True</property>
        <property name="ItemTemplateId" type="string">~sitecollection/_catalogs/masterpage/Display Templates/Content Web Parts/Item_KPCU_Happenings.js</property>
        <property name="PropertyMappings" type="string" />
        <property name="ChromeState" type="chromestate">Normal</property>
        <property name="IncludeResultTypeConstraint" type="bool">False</property>
        <property name="StartingItemIndex" type="int">1</property>
        <property name="ShowDefinitions" type="bool">False</property>
        <property name="Height" type="string" />
        <property name="Hidden" type="bool">False</property>
        <property name="HitHighlightedPropertiesJson" type="string">["Title","Path","Author","SectionNames","SiteDescription"]</property>
        <property name="ScrollToTopOnRedraw" type="bool">False</property>
        <property name="UseSharedDataProvider" type="bool">False</property>
        <property name="RepositionLanguageDropDown" type="bool">False</property>
        <property name="AlwaysRenderOnServer" type="bool">False</property>
        <property name="AllowConnect" type="bool">True</property>
        <property name="ItemBodyTemplateId" type="string" />
        <property name="ShowAlertMe" type="bool">True</property>
        <property name="ExportMode" type="exportmode">All</property>
        <property name="AddSEOPropertiesFromSearch" type="bool">False</property>
        <property name="ShowUpScopeMessage" type="bool">False</property>
        <property name="AllowHide" type="bool">True</property>
        <property name="AllowClose" type="bool">True</property>
        <property name="UseSimplifiedQueryBuilder" type="bool">False</property>
        <property name="ShouldHideControlWhenEmpty" type="bool">False</property>
        <property name="ResultType" type="string" />
        <property name="LogAnalyticsViewEvent" type="bool">False</property>
        <property name="MaxPagesAfterCurrent" type="int">1</property>
        <property name="TitleUrl" type="string" />
        <property name="EmptyMessage" type="string" />
        <property name="AdvancedSearchPageAddress" type="string">advanced.aspx</property>
        <property name="IsXGeo3SForwardingFlighted" type="bool">True</property>
        <property name="AllowMinimize" type="bool">True</property>
        <property name="ShowBestBets" type="bool">False</property>
        <property name="AllowEdit" type="bool">True</property>
        <property name="NumberOfItems" type="int">4</property>
        <property name="HelpUrl" type="string" />
        <property name="ShowPaging" type="bool">True</property>
        <property name="ShowViewDuplicates" type="bool">False</property>
        <property name="SelectedPropertiesJson" type="string">["EventType","owstaxIdEventType1","Path","Title","EventDateOWSDATE","EndDateOWSDATE","MPLocation","SecondaryFileExtension","ContentTypeId"]</property>
        <property name="TargetResultTable" type="string">RelevantResults</property>
        <property name="HelpMode" type="helpmode">Modeless</property>
        <property name="ShowXGeoOptions" type="bool">False</property>
        <property name="IsXGeoFlighted" type="bool">False</property>
        <property name="ShowPersonalFavorites" type="bool">False</property>
        <property name="EnableXGeo3SForwarding" type="bool">False</property>
        <property name="PreloadedItemTemplateIdsJson" type="string">null</property>
        <property name="Description" type="string">Content Search Web Part will allow you to show items that are results of a search query you specify. When you add it to the page, this Web Part will show recently modified items from the current site. You can change this setting to show items from another site or list by editing the Web Part and changing its search criteria.As new content is discovered by search, this Web Part will display an updated list of items each time the page is viewed.</property>
        <property name="ShowPreferencesLink" type="bool">True</property>
        <property name="QueryGroupName" type="string">6f9fcc19-aeaa-42f9-9b57-34f0dcd8b8af</property>
        <property name="ShowResultCount" type="bool">True</property>
        <property name="TitleIconImageUrl" type="string" />
        <property name="Direction" type="direction">NotSet</property>
        <property name="ResultsPerPage" type="int">4</property>
        <property name="AvailableSortsJson" type="string">null</property>
        <property name="ShowResults" type="bool">True</property>
        <property name="ServerIncludeScriptsJson" type="string">null</property>
        <property name="SearchCenterXGeoLocations" type="string" />
        <property name="DataProviderJSON" type="string">{"QueryGroupName":"6f9fcc19-aeaa-42f9-9b57-34f0dcd8b8af","QueryPropertiesTemplateUrl":"sitesearch://webroot","IgnoreQueryPropertiesTemplateUrl":false,"SourceID":"8413cd39-2156-4e00-b54d-11efd9abdb89","SourceName":"Local SharePoint Results","SourceLevel":"Ssa","CollapseSpecification":"","QueryTemplate":"path:\"https://kpcu.sharepoint.com/sites/Intranet-Dev/Lists/Happenings/\"  ContentType:KPCU_Happenings RefinableDate01\u003e={Today}","FallbackSort":[{"p":"RefinableDate01","d":0}],"FallbackSortJson":"[{\"p\":\"RefinableDate01\",\"d\":0}]","RankRules":null,"RankRulesJson":"null","AsynchronousResultRetrieval":false,"SendContentBeforeQuery":true,"BatchClientQuery":true,"FallbackLanguage":-1,"FallbackRankingModelID":"","EnableStemming":true,"EnablePhonetic":false,"EnableNicknames":false,"EnableInterleaving":false,"EnableQueryRules":true,"EnableOrderingHitHighlightedProperty":false,"HitHighlightedMultivaluePropertyLimit":-1,"IgnoreContextualScope":true,"ScopeResultsToCurrentSite":false,"TrimDuplicates":false,"Properties":{"TryCache":true,"Scope":"{Site.URL}","UpdateLinksForCatalogItems":true,"EnableStacking":true,"CrossGeoQuery":"false","ListId":"55e27111-388b-411d-87df-351f2b92600f","ListItemId":4},"PropertiesJson":"{\"TryCache\":true,\"Scope\":\"{Site.URL}\",\"UpdateLinksForCatalogItems\":true,\"EnableStacking\":true,\"CrossGeoQuery\":\"false\",\"ListId\":\"55e27111-388b-411d-87df-351f2b92600f\",\"ListItemId\":4}","ClientType":"ContentSearchRegular","ClientFunction":"","ClientFunctionDetails":"","UpdateAjaxNavigate":true,"SummaryLength":180,"DesiredSnippetLength":90,"PersonalizedQuery":false,"FallbackRefinementFilters":null,"IgnoreStaleServerQuery":false,"RenderTemplateId":"DefaultDataProvider","AlternateErrorMessage":null,"Title":""}</property>
        <property name="ShowAdvancedLink" type="bool">True</property>
        <property name="ShowDidYouMean" type="bool">False</property>
        <property name="AllowZoneChange" type="bool">True</property>
        <property name="ChromeType" type="chrometype">None</property>
        <property name="GroupTemplateId" type="string">~sitecollection/_catalogs/masterpage/Display Templates/Content Web Parts/Group_Content.js</property>
        <property name="MissingAssembly" type="string">Cannot import this Web Part.</property>
        <property name="OverwriteResultPath" type="bool">True</property>
        <property name="Width" type="string" />
        <property name="MaxPagesBeforeCurrent" type="int">4</property>
        <property name="XGeoTenantsInfo" type="string" />
        <property name="ShowLanguageOptions" type="bool">True</property>
        <property name="ResultTypeId" type="string" />
        <property name="AlternateErrorMessage" type="string" null="true" />
        <property name="Title" type="string">Happenings</property>
        <property name="RenderTemplateId" type="string">~sitecollection/_catalogs/masterpage/Display Templates/Content Web Parts/control_Happenings.js</property>
        <property name="EmitStyleReference" type="bool">True</property>
        <property name="StatesJson" type="string">{}</property>
        <property name="ShowSortOptions" type="bool">False</property>
        <property name="CatalogIconImageUrl" type="string" />
      </properties>
    </data>
  </webPart>
</webParts>`;

    var voiceOfManagementXml = `<webParts>
  <webPart xmlns="http://schemas.microsoft.com/WebPart/v3">
    <metaData>
      <type name="Microsoft.Office.Server.Search.WebControls.ContentBySearchWebPart, Microsoft.Office.Server.Search, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" />
      <importErrorMessage>Cannot import this Web Part.</importErrorMessage>
    </metaData>
    <data>
      <properties>
        <property name="BypassResultTypes" type="bool">True</property>
        <property name="ItemTemplateId" type="string">~sitecollection/_catalogs/masterpage/Display Templates/Content Web Parts/Item_VoiceOfManagement.js</property>
        <property name="PropertyMappings" type="string" />
        <property name="ChromeState" type="chromestate">Normal</property>
        <property name="IncludeResultTypeConstraint" type="bool">False</property>
        <property name="StartingItemIndex" type="int">1</property>
        <property name="ShowDefinitions" type="bool">False</property>
        <property name="Height" type="string" />
        <property name="Hidden" type="bool">False</property>
        <property name="HitHighlightedPropertiesJson" type="string">["Title","Path","Author","SectionNames","SiteDescription"]</property>
        <property name="ScrollToTopOnRedraw" type="bool">False</property>
        <property name="UseSharedDataProvider" type="bool">False</property>
        <property name="RepositionLanguageDropDown" type="bool">False</property>
        <property name="AlwaysRenderOnServer" type="bool">False</property>
        <property name="AllowConnect" type="bool">True</property>
        <property name="ItemBodyTemplateId" type="string" />
        <property name="ShowAlertMe" type="bool">False</property>
        <property name="ExportMode" type="exportmode">All</property>
        <property name="AddSEOPropertiesFromSearch" type="bool">False</property>
        <property name="ShowUpScopeMessage" type="bool">False</property>
        <property name="AllowHide" type="bool">True</property>
        <property name="AllowClose" type="bool">True</property>
        <property name="UseSimplifiedQueryBuilder" type="bool">False</property>
        <property name="ShouldHideControlWhenEmpty" type="bool">False</property>
        <property name="ResultType" type="string" />
        <property name="LogAnalyticsViewEvent" type="bool">False</property>
        <property name="MaxPagesAfterCurrent" type="int">1</property>
        <property name="TitleUrl" type="string" />
        <property name="EmptyMessage" type="string" />
        <property name="AdvancedSearchPageAddress" type="string">advanced.aspx</property>
        <property name="IsXGeo3SForwardingFlighted" type="bool">True</property>
        <property name="AllowMinimize" type="bool">True</property>
        <property name="ShowBestBets" type="bool">False</property>
        <property name="AllowEdit" type="bool">True</property>
        <property name="NumberOfItems" type="int">1</property>
        <property name="HelpUrl" type="string" />
        <property name="ShowPaging" type="bool">True</property>
        <property name="ShowViewDuplicates" type="bool">False</property>
        <property name="SelectedPropertiesJson" type="string">["ApprovalStatus","PublishingImage","PictureURL","PictureThumbnailURL","Path","Title","ManagementMessage","ContentAuthor","SecondaryFileExtension","ContentTypeId"]</property>
        <property name="TargetResultTable" type="string">RelevantResults</property>
        <property name="HelpMode" type="helpmode">Modeless</property>
        <property name="ShowXGeoOptions" type="bool">False</property>
        <property name="IsXGeoFlighted" type="bool">False</property>
        <property name="ShowPersonalFavorites" type="bool">False</property>
        <property name="EnableXGeo3SForwarding" type="bool">False</property>
        <property name="PreloadedItemTemplateIdsJson" type="string">null</property>
        <property name="Description" type="string">Content Search Web Part will allow you to show items that are results of a search query you specify. When you add it to the page, this Web Part will show recently modified items from the current site. You can change this setting to show items from another site or list by editing the Web Part and changing its search criteria.As new content is discovered by search, this Web Part will display an updated list of items each time the page is viewed.</property>
        <property name="ShowPreferencesLink" type="bool">True</property>
        <property name="QueryGroupName" type="string">8a80b56d-84db-4ece-8a90-07a13d352681</property>
        <property name="ShowResultCount" type="bool">True</property>
        <property name="TitleIconImageUrl" type="string" />
        <property name="Direction" type="direction">NotSet</property>
        <property name="ResultsPerPage" type="int">1</property>
        <property name="AvailableSortsJson" type="string">null</property>
        <property name="ShowResults" type="bool">True</property>
        <property name="ServerIncludeScriptsJson" type="string">null</property>
        <property name="SearchCenterXGeoLocations" type="string" />
        <property name="DataProviderJSON" type="string">{"QueryGroupName":"8a80b56d-84db-4ece-8a90-07a13d352681","QueryPropertiesTemplateUrl":"sitesearch://webroot","IgnoreQueryPropertiesTemplateUrl":false,"SourceID":"8413cd39-2156-4e00-b54d-11efd9abdb89","SourceName":"Local SharePoint Results","SourceLevel":"Ssa","CollapseSpecification":"","QueryTemplate":"(contentclass:STS_ListItem OR IsDocument:True) ContentType:\"Voice_of_Management\" Path:getSearchItemsPath","FallbackSort":[{"p":"RefinableString00","d":0},{"p":"LastModifiedTime","d":0}],"FallbackSortJson":"[{\"p\":\"RefinableString00\",\"d\":0},{\"p\":\"LastModifiedTime\",\"d\":0}]","RankRules":null,"RankRulesJson":"null","AsynchronousResultRetrieval":false,"SendContentBeforeQuery":true,"BatchClientQuery":true,"FallbackLanguage":-1,"FallbackRankingModelID":"","EnableStemming":true,"EnablePhonetic":false,"EnableNicknames":false,"EnableInterleaving":false,"EnableQueryRules":true,"EnableOrderingHitHighlightedProperty":false,"HitHighlightedMultivaluePropertyLimit":-1,"IgnoreContextualScope":true,"ScopeResultsToCurrentSite":false,"TrimDuplicates":false,"Properties":{"TryCache":true,"Scope":"{Site.URL}","UpdateLinksForCatalogItems":true,"EnableStacking":true,"ListId":"3b20d6ec-6bba-4b5d-903e-6d1ab7febb71","ListItemId":5,"CrossGeoQuery":"false"},"PropertiesJson":"{\"TryCache\":true,\"Scope\":\"{Site.URL}\",\"UpdateLinksForCatalogItems\":true,\"EnableStacking\":true,\"ListId\":\"3b20d6ec-6bba-4b5d-903e-6d1ab7febb71\",\"ListItemId\":5,\"CrossGeoQuery\":\"false\"}","ClientType":"ContentSearchRegular","ClientFunction":"","ClientFunctionDetails":"","UpdateAjaxNavigate":true,"SummaryLength":180,"DesiredSnippetLength":90,"PersonalizedQuery":false,"FallbackRefinementFilters":null,"IgnoreStaleServerQuery":false,"RenderTemplateId":"DefaultDataProvider","AlternateErrorMessage":null,"Title":""}</property>
        <property name="ShowAdvancedLink" type="bool">True</property>
        <property name="ShowDidYouMean" type="bool">False</property>
        <property name="AllowZoneChange" type="bool">True</property>
        <property name="ChromeType" type="chrometype">None</property>
        <property name="GroupTemplateId" type="string">~sitecollection/_catalogs/masterpage/Display Templates/Content Web Parts/Group_Content.js</property>
        <property name="MissingAssembly" type="string">Cannot import this Web Part.</property>
        <property name="OverwriteResultPath" type="bool">True</property>
        <property name="Width" type="string" />
        <property name="MaxPagesBeforeCurrent" type="int">4</property>
        <property name="XGeoTenantsInfo" type="string" />
        <property name="ShowLanguageOptions" type="bool">True</property>
        <property name="ResultTypeId" type="string" />
        <property name="AlternateErrorMessage" type="string" null="true" />
        <property name="Title" type="string">Voice of Mgt CSWP</property>
        <property name="RenderTemplateId" type="string">~sitecollection/_catalogs/masterpage/Display Templates/Content Web Parts/Control_VoiceOfMgtPaging.js</property>
        <property name="EmitStyleReference" type="bool">True</property>
        <property name="StatesJson" type="string">{}</property>
        <property name="ShowSortOptions" type="bool">False</property>
        <property name="CatalogIconImageUrl" type="string" />
      </properties>
    </data>
  </webPart>
</webParts>`;

    var latestUpdatesWebpartXml = `<webParts>
  <webPart xmlns="http://schemas.microsoft.com/WebPart/v3">
    <metaData>
      <type name="Microsoft.Office.Server.Search.WebControls.ContentBySearchWebPart, Microsoft.Office.Server.Search, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" />
      <importErrorMessage>Cannot import this Web Part.</importErrorMessage>
    </metaData>
    <data>
      <properties>
        <property name="BypassResultTypes" type="bool">True</property>
        <property name="ItemTemplateId" type="string">~sitecollection/_catalogs/masterpage/Display Templates/Content Web Parts/Item_LatestUpdateHomePage.js</property>
        <property name="PropertyMappings" type="string" />
        <property name="ChromeState" type="chromestate">Normal</property>
        <property name="IncludeResultTypeConstraint" type="bool">False</property>
        <property name="StartingItemIndex" type="int">1</property>
        <property name="ShowDefinitions" type="bool">False</property>
        <property name="Height" type="string" />
        <property name="Hidden" type="bool">False</property>
        <property name="HitHighlightedPropertiesJson" type="string">["Title","Path","Author","SectionNames","SiteDescription"]</property>
        <property name="ScrollToTopOnRedraw" type="bool">False</property>
        <property name="UseSharedDataProvider" type="bool">False</property>
        <property name="RepositionLanguageDropDown" type="bool">False</property>
        <property name="AlwaysRenderOnServer" type="bool">False</property>
        <property name="AllowConnect" type="bool">True</property>
        <property name="ItemBodyTemplateId" type="string" />
        <property name="ShowAlertMe" type="bool">False</property>
        <property name="ExportMode" type="exportmode">All</property>
        <property name="AddSEOPropertiesFromSearch" type="bool">False</property>
        <property name="ShowUpScopeMessage" type="bool">False</property>
        <property name="AllowHide" type="bool">True</property>
        <property name="AllowClose" type="bool">True</property>
        <property name="UseSimplifiedQueryBuilder" type="bool">False</property>
        <property name="ShouldHideControlWhenEmpty" type="bool">False</property>
        <property name="ResultType" type="string" />
        <property name="LogAnalyticsViewEvent" type="bool">False</property>
        <property name="MaxPagesAfterCurrent" type="int">1</property>
        <property name="TitleUrl" type="string" />
        <property name="EmptyMessage" type="string" />
        <property name="AdvancedSearchPageAddress" type="string">advanced.aspx</property>
        <property name="IsXGeo3SForwardingFlighted" type="bool">True</property>
        <property name="AllowMinimize" type="bool">True</property>
        <property name="ShowBestBets" type="bool">False</property>
        <property name="AllowEdit" type="bool">True</property>
        <property name="NumberOfItems" type="int">1</property>
        <property name="HelpUrl" type="string" />
        <property name="ShowPaging" type="bool">True</property>
        <property name="ShowViewDuplicates" type="bool">False</property>
        <property name="SelectedPropertiesJson" type="string">["ListItemID","StoryCopy","BannerImageOWSIMGE","PublishingImage","PictureURL","PictureThumbnailURL","Path","Title","EmployeeDepartment","BannerImage","SecondaryFileExtension","ContentTypeId"]</property>
        <property name="TargetResultTable" type="string">RelevantResults</property>
        <property name="HelpMode" type="helpmode">Modeless</property>
        <property name="ShowXGeoOptions" type="bool">False</property>
        <property name="IsXGeoFlighted" type="bool">False</property>
        <property name="ShowPersonalFavorites" type="bool">False</property>
        <property name="EnableXGeo3SForwarding" type="bool">False</property>
        <property name="PreloadedItemTemplateIdsJson" type="string">null</property>
        <property name="Description" type="string">Content Search Web Part will allow you to show items that are results of a search query you specify. When you add it to the page, this Web Part will show recently modified items from the current site. You can change this setting to show items from another site or list by editing the Web Part and changing its search criteria.As new content is discovered by search, this Web Part will display an updated list of items each time the page is viewed.</property>
        <property name="ShowPreferencesLink" type="bool">True</property>
        <property name="QueryGroupName" type="string">35afa177-8ad2-485c-9f84-3eab99ab28c3</property>
        <property name="ShowResultCount" type="bool">True</property>
        <property name="TitleIconImageUrl" type="string" />
        <property name="Direction" type="direction">NotSet</property>
        <property name="ResultsPerPage" type="int">1</property>
        <property name="AvailableSortsJson" type="string">null</property>
        <property name="ShowResults" type="bool">True</property>
        <property name="ServerIncludeScriptsJson" type="string">null</property>
        <property name="SearchCenterXGeoLocations" type="string" />
        <property name="DataProviderJSON" type="string">{"QueryGroupName":"35afa177-8ad2-485c-9f84-3eab99ab28c3","QueryPropertiesTemplateUrl":"sitesearch://webroot","IgnoreQueryPropertiesTemplateUrl":false,"SourceID":"8413cd39-2156-4e00-b54d-11efd9abdb89","SourceName":"Local SharePoint Results","SourceLevel":"Ssa","CollapseSpecification":"","QueryTemplate":"(contentclass:STS_ListItem OR IsDocument:True) ContentType:\"Latest_Updates\" Path:getSearchItemsPath","FallbackSort":[{"p":"RefinableString00","d":0},{"p":"LastModifiedTime","d":1}],"FallbackSortJson":"[{\"p\":\"RefinableString00\",\"d\":0},{\"p\":\"LastModifiedTime\",\"d\":1}]","RankRules":null,"RankRulesJson":"null","AsynchronousResultRetrieval":false,"SendContentBeforeQuery":true,"BatchClientQuery":true,"FallbackLanguage":-1,"FallbackRankingModelID":"","EnableStemming":true,"EnablePhonetic":false,"EnableNicknames":false,"EnableInterleaving":false,"EnableQueryRules":true,"EnableOrderingHitHighlightedProperty":false,"HitHighlightedMultivaluePropertyLimit":-1,"IgnoreContextualScope":true,"ScopeResultsToCurrentSite":false,"TrimDuplicates":false,"Properties":{"TryCache":true,"Scope":"{Site.URL}","UpdateLinksForCatalogItems":true,"EnableStacking":true,"ListId":"3b20d6ec-6bba-4b5d-903e-6d1ab7febb71","ListItemId":5,"CrossGeoQuery":"false"},"PropertiesJson":"{\"TryCache\":true,\"Scope\":\"{Site.URL}\",\"UpdateLinksForCatalogItems\":true,\"EnableStacking\":true,\"ListId\":\"3b20d6ec-6bba-4b5d-903e-6d1ab7febb71\",\"ListItemId\":5,\"CrossGeoQuery\":\"false\"}","ClientType":"ContentSearchRegular","ClientFunction":"","ClientFunctionDetails":"","UpdateAjaxNavigate":true,"SummaryLength":180,"DesiredSnippetLength":90,"PersonalizedQuery":false,"FallbackRefinementFilters":null,"IgnoreStaleServerQuery":false,"RenderTemplateId":"DefaultDataProvider","AlternateErrorMessage":null,"Title":""}</property>
        <property name="ShowAdvancedLink" type="bool">True</property>
        <property name="ShowDidYouMean" type="bool">False</property>
        <property name="AllowZoneChange" type="bool">True</property>
        <property name="ChromeType" type="chrometype">None</property>
        <property name="GroupTemplateId" type="string">~sitecollection/_catalogs/masterpage/Display Templates/Content Web Parts/Group_Content.js</property>
        <property name="MissingAssembly" type="string">Cannot import this Web Part.</property>
        <property name="OverwriteResultPath" type="bool">True</property>
        <property name="Width" type="string" />
        <property name="MaxPagesBeforeCurrent" type="int">4</property>
        <property name="XGeoTenantsInfo" type="string" />
        <property name="ShowLanguageOptions" type="bool">True</property>
        <property name="ResultTypeId" type="string" />
        <property name="AlternateErrorMessage" type="string" null="true" />
        <property name="Title" type="string">Latest Updates Home CSWP</property>
        <property name="RenderTemplateId" type="string">~sitecollection/_catalogs/masterpage/Display Templates/Content Web Parts/Control_LatestUpdates.js</property>
        <property name="EmitStyleReference" type="bool">True</property>
        <property name="StatesJson" type="string">{}</property>
        <property name="ShowSortOptions" type="bool">False</property>
        <property name="CatalogIconImageUrl" type="string" />
      </properties>
    </data>
  </webPart>
</webParts>`;

    var hostWebRelativeUrl = hostWeb.get_serverRelativeUrl();

    var latestUpdatesWebPart = limitedWebPartManager.importWebPart(latestUpdatesWebpartXml).get_webPart();

    var voiceOfManagement = limitedWebPartManager.importWebPart(voiceOfManagementXml).get_webPart();

    var eventsWebPart = limitedWebPartManager.importWebPart(eventsWebPartXml).get_webPart();

    var orgAnnouncementWebPart = limitedWebPartManager.importWebPart(orgAnnouncementWebPartXml.replace("{site_relative_url}", hostWebRelativeUrl)).get_webPart();

    var newsWebPart = limitedWebPartManager.importWebPart(newsWebPartXml).get_webPart();

    var pollWebPart = limitedWebPartManager.importWebPart(pollWebPartXml.replace("{site_relative_url}", hostWebRelativeUrl)).get_webPart();

    var birthDayWebPart = limitedWebPartManager.importWebPart(birthDayWebPartXml.replace("{site_relative_url}", hostWebRelativeUrl)).get_webPart();

    var zoneID = "xeb4fdf5f1db747c0b7e47d6b2af8a19f";
    limitedWebPartManager.addWebPart(newsWebPart, zoneID, 1);
    limitedWebPartManager.addWebPart(orgAnnouncementWebPart, zoneID, 2);
    limitedWebPartManager.addWebPart(pollWebPart, zoneID, 3);
    limitedWebPartManager.addWebPart(birthDayWebPart, zoneID, 4);
    limitedWebPartManager.addWebPart(eventsWebPart, zoneID, 5);
    limitedWebPartManager.addWebPart(voiceOfManagement, zoneID, 6);
    limitedWebPartManager.addWebPart(latestUpdatesWebPart, zoneID, 7);

    oFile.checkIn();
    oFile.publish();
    context.load(newsWebPart);
    context.load(orgAnnouncementWebPart);
    context.load(pollWebPart);
    context.load(birthDayWebPart);
    context.load(eventsWebPart);
    context.load(voiceOfManagement);
    context.load(latestUpdatesWebPart);

    context.executeQueryAsync(Function.createDelegate(this, function onQuerySucceeded() {
        alert('Web Part added: ' + newsWebPart.get_title());
    }), Function.createDelegate(this, function onQueryFailed(sender, args) {
        alert('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
    }));
}