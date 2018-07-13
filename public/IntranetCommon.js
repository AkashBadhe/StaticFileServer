'use strict';
var INTRANET = INTRANET || {};

INTRANET.Constant = {
    SiteColumnGroupName: 'IntranetInbox'
}
INTRANET.Schema = {};

INTRANET.LogType = {
    Info: 'Information',
    Error: 'Error'
};

INTRANET.Schema.SiteColumns = [
    '<Field Command= "Create" Type= "Choice" DisplayName= "Email Notification" Required= "FALSE" EnforceUniqueValues= "FALSE" Indexed= "FALSE" MaxLength= "255" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName= "EmailNotification" Name= "EmailNotification">' +
    '<CHOICES>' +
    '<CHOICE>Yes</CHOICE>' +
    '<CHOICE>No</CHOICE>' +
    '</CHOICES>' +
    '</Field> ',
    '<Field Command="Create" Type="DateTime" DisplayName="End Date1" Required="TRUE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" Group="' + INTRANET.Constant.SiteColumnGroupName + '" StaticName="EndDate1" Name="EndDate1"></Field>',
    '<Field Command= "Create" Type= "Choice" DisplayName= "Is Poll" Required= "TRUE" EnforceUniqueValues= "FALSE" Indexed= "FALSE" MaxLength= "255" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName= "IsPoll" Name= "IsPoll">' +
    '<CHOICES>' +
    '<CHOICE>Yes</CHOICE>' +
    '<CHOICE>No</CHOICE>' +
    '</CHOICES>' +
    '</Field> ',
    '<Field Command= "Create" Type= "URL" DisplayName= "Manage Questions" Required= "TRUE" EnforceUniqueValues= "FALSE" Indexed= "FALSE" Sortable= "FALSE" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName= "ManageQuestions" Name= "ManageQuestions" ></Field >',
    '<Field Command= "Create" Type= "URL" DisplayName= "View Response" Required= "TRUE" EnforceUniqueValues= "FALSE" Indexed= "FALSE" Sortable= "FALSE" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName= "ViewResponse" Name= "ViewResponse" ></Field >',
    '<Field Command= "Create" Type= "URL" DisplayName= "Survey Link" Required= "TRUE" EnforceUniqueValues= "FALSE" Indexed= "FALSE" Sortable= "FALSE" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName= "SurveyLink" Name= "SurveyLink" ></Field >',
    '<Field Command= "Create" Type= "Note" DisplayName= "Survey Description" Required= "TRUE" EnforceUniqueValues= "FALSE" Indexed= "FALSE" NumLines= "6" RichText= "FALSE" Sortable= "FALSE" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName= "SurveyDescription" Name= "SurveyDescription" ></Field >',
    '<Field Command= "Create" Type= "Choice" DisplayName= "Content Priority" Required= "TRUE" EnforceUniqueValues= "FALSE" Indexed= "FALSE" MaxLength= "255" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName= "ContentPriority" Name= "ContentPriority">' +
    '<CHOICES>' +
    '<CHOICE>1</CHOICE>' +
    '<CHOICE>2</CHOICE>' +
    '<CHOICE>3</CHOICE>' +
    '<CHOICE>4</CHOICE>' +
    '<CHOICE>5</CHOICE>' +
    '</CHOICES>' +
    '</Field> ',
    '<Field Command= "Create" Type= "Note" DisplayName= "News Details" Required= "TRUE" EnforceUniqueValues= "FALSE" Indexed= "FALSE" NumLines= "6" RichText= "TRUE" Sortable= "FALSE" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName= "NewsDetails" Name= "NewsDetails"></Field >',
    '<Field Command= "Create" Type= "User" DisplayName= "Content Author" List= "UserInfo" Required= "FALSE" EnforceUniqueValues= "FALSE" ShowField= "ImnName" UserSelectionMode= "PeopleOnly" UserSelectionScope= "0" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName= "ContentAuthor" Name= "ContentAuthor"></Field >',
    '<Field Command= "Create" Type= "User" DisplayName= "EmployeeName" List= "UserInfo" Required= "FALSE" EnforceUniqueValues= "FALSE" ShowField= "ImnName" UserSelectionMode= "PeopleOnly" UserSelectionScope= "0" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName= "EmployeeName" Name= "EmployeeName"></Field >',
    '<Field Command= "Create" Type= "Note" DisplayName= "StoryCopy" Required= "TRUE" EnforceUniqueValues= "FALSE" Indexed= "FALSE" NumLines= "6" RichText= "TRUE" Sortable= "FALSE" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName= "StoryCopy" Name= "StoryCopy" ></Field >',
    '<Field Command= "Create" Type= "Note" DisplayName= "Announcement Details" Required= "TRUE" EnforceUniqueValues= "FALSE" Indexed= "FALSE" NumLines= "6" RichText= "TRUE" Sortable= "FALSE" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName= "AnnouncementDetails" Name= "AnnouncementDetails" ></Field >',
    '<Field Command= "Create" Type= "Note" DisplayName= "Management Message" Required= "TRUE" EnforceUniqueValues= "FALSE" Indexed= "FALSE" NumLines= "6" RichText= "TRUE" Sortable= "FALSE" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName= "ManagementMessage" Name= "ManagementMessage" ></Field >',
    '<Field Command= "Create" Type="Image" DisplayName= "Banner Image" Required= "TRUE" RichText="TRUE" RichTextMode="ThemeHtml" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName="BannerImage" Name="BannerImage"></Field >',
    '<Field Command= "Create" Type="Image" DisplayName= "Home Thumbnail" Required= "TRUE" RichText="TRUE" RichTextMode="ThemeHtml" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName="HomeThumbnail" Name="HomeThumbnail"></Field >',
    '<Field Command= "Create" Type="Image" DisplayName= "Thumbnail Image" Required= "TRUE" RichText="TRUE" RichTextMode="ThemeHtml" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName="ThumbnailImage" Name="ThumbnailImage"></Field >',
    '<Field Command= "Create" Type= "Choice" DisplayName= "Audit Type" Required= "TRUE" EnforceUniqueValues= "FALSE" Indexed= "FALSE" MaxLength= "255" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName= "AuditType" Name= "AuditType">' +
    '<CHOICES>' +
    '<CHOICE>Planned</CHOICE>' +
    '<CHOICE>Special</CHOICE>' +
    '</CHOICES>' +
    '</Field> ',
    '<Field Command= "Create" Type= "Choice" DisplayName= "Audit Category" Required= "TRUE" EnforceUniqueValues= "FALSE" Indexed= "FALSE" MaxLength= "255" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName= "AuditCategory" Name= "AuditCategory">' +
    '<CHOICES>' +
    '<CHOICE>Credit</CHOICE>' +
    '<CHOICE>Operations</CHOICE>' +
    '<CHOICE>IT</CHOICE>' +
    '<CHOICE>Support Functions</CHOICE>' +
    '<CHOICE>Domestic Branches</CHOICE>' +
    '<CHOICE>Overseas Branches</CHOICE>' +
    '<CHOICE>Other</CHOICE>' +
    '</CHOICES>' +
    '</Field> ',
    '<Field Command= "Create" Type= "Choice" DisplayName= "Issue Significance" Required= "TRUE" EnforceUniqueValues= "FALSE" Indexed= "FALSE" MaxLength= "255" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName= "IssueSignificance" Name= "IssueSignificance">' +
    '<CHOICES>' +
    '<CHOICE>High</CHOICE>' +
    '<CHOICE>Medium</CHOICE>' +
    '<CHOICE>Low</CHOICE>' +
    '</CHOICES>' +
    '</Field> ',
    '<Field Command= "Create" Type= "Choice" DisplayName= "Issue Status" Required= "TRUE" EnforceUniqueValues= "FALSE" Indexed= "FALSE" MaxLength= "255" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName= "IssueStatus" Name= "IssueStatus">' +
    '<CHOICES>' +
    '<CHOICE>High</CHOICE>' +
    '<CHOICE>Medium</CHOICE>' +
    '<CHOICE>Low</CHOICE>' +
    '</CHOICES>' +
    '</Field> ',
    '<Field Command= "Create" Type="Text" DisplayName= "Audit Title" Required= "FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName="AuditTitle" Name="AuditTitle"></Field >',
    '<Field Command= "Create" Type="Text" DisplayName= "Audit Assignment Notice No" Required= "FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName="AuditAssignmentNoticeNo" Name="AuditAssignmentNoticeNo"></Field >',
    '<Field Command= "Create" Type="Text" DisplayName= "ParentIssueID" Required= "FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName="ParentIssueID" Name="ParentIssueID"></Field >',
    '<Field Command= "Create" Type="Text" DisplayName= "ParentAuditTitle" Required= "FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName="ParentAuditTitle" Name="ParentAuditTitle"></Field >',
    '<Field Command= "Create" Type="DateTime" DisplayName= "Original Due Date" Required= "FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName="OriginalDueDate" Name="OriginalDueDate"></Field >',
    '<Field Command= "Create" Type="DateTime" DisplayName= "Audit Report Date" Required= "FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName="AuditReportDate" Name="AuditReportDate"></Field >',
    '<Field Command= "Create" Type="DateTime" DisplayName= "Revised Due Date" Required= "FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName="RevisedDueDate" Name="RevisedDueDate"></Field >',
    '<Field Command= "Create" Type="DateTime" DisplayName= "Issue Closure Date" Required= "FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName="IssueClosureDate" Name="IssueClosureDate"></Field >',
    '<Field Command= "Create" Type="Note" DisplayName= "Audit Issue Summary" Required= "FALSE" NumLines= "6" RichText= "TRUE" Sortable= "FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName="AuditIssueSummary" Name="AuditIssueSummary"></Field >',
    '<Field Command= "Create" Type="Note" DisplayName= "Issue Details" Required= "FALSE" NumLines= "6" RichText= "TRUE" Sortable= "FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName="IssueDetails" Name="IssueDetails"></Field >',
    '<Field Command= "Create" Type="Note" DisplayName= "Audit Recommendation" Required= "FALSE" NumLines= "6" RichText= "TRUE" Sortable= "FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName="AuditRecommendation" Name="AuditRecommendation"></Field >',
    '<Field Command= "Create" Type="Note" DisplayName= "Agreed Corrective Action" Required= "FALSE" NumLines= "6" RichText= "TRUE" Sortable= "FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName="AgreedCorrectiveAction" Name="AgreedCorrectiveAction"></Field >',
    '<Field Command= "Create" Type="Note" DisplayName= "Issue Resolution Comments" Required= "FALSE" NumLines= "6" RichText= "TRUE" Sortable= "FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName="IssueResolutionComments" Name="IssueResolutionComments"></Field >',
    '<Field Command= "Create" Type="Note" DisplayName= "Auditor Remarks" Required= "FALSE" NumLines= "6" RichText= "TRUE" Sortable= "FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName="AuditorRemarks" Name="AuditorRemarks"></Field >',
    '<Field Command= "Create" Type= "User" DisplayName= "Person to Address(PTA)" List= "UserInfo" Required= "FALSE" EnforceUniqueValues= "FALSE" ShowField= "ImnName" UserSelectionMode= "PeopleAndGroups" UserSelectionScope= "0" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName= "PersonToAddress" Name= "PersonToAddress"></Field >',
    '<Field Command= "Create" Type= "User" DisplayName= "Supervisory Level" List= "UserInfo" Required= "FALSE" EnforceUniqueValues= "FALSE" ShowField= "ImnName" UserSelectionMode= "PeopleAndGroups" UserSelectionScope= "0" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName= "SupervisoryLevel" Name= "SupervisoryLevel"></Field >',
    '<Field Command= "Create" Type= "User" DisplayName= "Auditor(s)" List= "UserInfo" Required= "FALSE" EnforceUniqueValues= "FALSE" ShowField= "ImnName" UserSelectionMode= "PeopleAndGroups" UserSelectionScope= "0" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName= "Auditors" Name= "Auditors"></Field >',
    '<Field Command= "Create" Type= "User" DisplayName= "AuditorManager(s)" List= "UserInfo" Required= "FALSE" EnforceUniqueValues= "FALSE" ShowField= "ImnName" UserSelectionMode= "PeopleAndGroups" UserSelectionScope= "0" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName= "AuditManagers" Name= "AuditManagers"></Field >',
    '<Field Command= "Create" Type= "User" DisplayName= "Issue Owner" List= "UserInfo" Required= "FALSE" EnforceUniqueValues= "FALSE" ShowField= "ImnName" UserSelectionMode= "PeopleAndGroups" UserSelectionScope= "0" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName= "IssueOwner" Name= "IssueOwner"></Field >',
    '<Field Command= "Create" Type="Calculated" DisplayName= "Due Status"  Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName="DueStatus" Name="DueStatus">' +
    '<FieldRefs>' +
    '<FieldRef Name="OriginalDueDate"/>' +
    '</FieldRefs>' +
    '<Formula>=[Original Due Date]</Formula>' +
    '</Field >',
    '<Field Command= "Create" Type="Text" DisplayName= "Key" Required= "TRUE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName="Key" Name="Key"></Field >',
    '<Field Command= "Create" Type="Text" DisplayName= "Value" Required= "TRUE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName="Value" Name="Value"></Field >',
    '<Field Command= "Create" Type="Text" DisplayName= "Category" Required= "TRUE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName="Category" Name="Category"></Field >',
    '<Field Command= "Create" Type= "Choice" DisplayName= "NavigationLevel" Required= "FALSE" EnforceUniqueValues= "FALSE" Indexed= "FALSE" MaxLength= "255" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName= "NavigationLevel" Name= "NavigationLevel">' +
    '<CHOICES>' +
    '<CHOICE>Level 1</CHOICE>' +
    '<CHOICE>Level 2</CHOICE>' +
    '<CHOICE>Level 3</CHOICE>' +
    '</CHOICES>' +
    '</Field> ',
    '<Field Command= "Create" Type= "Choice" DisplayName= "Open In New Tab" Required= "FALSE" EnforceUniqueValues= "FALSE" Indexed= "FALSE" MaxLength= "255" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName= "OpenInNewTab" Name= "OpenInNewTab">' +
    '<CHOICES>' +
    '<CHOICE>Yes</CHOICE>' +
    '<CHOICE>No</CHOICE>' +
    '</CHOICES>' +
    '</Field> ',
    '<Field Command= "Create" Type= "Choice" DisplayName= "IsActiveLink" Required= "FALSE" EnforceUniqueValues= "FALSE" Indexed= "FALSE" MaxLength= "255" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName= "IsActiveLink" Name= "IsActiveLink">' +
    '<CHOICES>' +
    '<CHOICE>Yes</CHOICE>' +
    '<CHOICE>No</CHOICE>' +
    '</CHOICES>' +
    '</Field> ',
    '<Field Name="Sequence" StaticName="Sequence" DisplayName="Sequence" Type="Number" Group= "' + INTRANET.Constant.SiteColumnGroupName + '"/>',
    '<Field Command= "Create" Type="Text" DisplayName= "NavigationLink" Required= "FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName="NavigationLink" Name="NavigationLink"></Field >',
    '<Field Command= "Create" Type="Text" DisplayName= "Parent" Required= "FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName="Parent" Name="Parent"></Field >',
    '<Field Type="Text" DisplayName="News Title" Required="TRUE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName="NewsTitle" Name="NewsTitle" CustomFormatter=""></Field >',
    '<Field Type="DateTime" DisplayName="News Published Date" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateOnly" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" FriendlyDisplayFormat="Disabled" StaticName="NewsPublishedDate" Name="NewsPublishedDate" CustomFormatter="" CalType="0"></Field>',
    '<Field Type="User" DisplayName="News Author" List="UserInfo" Required="FALSE" EnforceUniqueValues="FALSE" ShowField="ImnName" UserSelectionMode="PeopleOnly" UserSelectionScope="0" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName="NewsAuthor" Name="NewsAuthor" />',
    '<Field Type="Image" DisplayName="News Banner" RichText="TRUE" RichTextMode="ThemeHtml" Required="FALSE" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" StaticName="NewsBanner" Name="NewsBanner" />',
    '<Field Name="PublishingPageContent" StaticName="PublishingPageContent" SourceID="http://schemas.microsoft.com/sharepoint/v3" Group= "' + INTRANET.Constant.SiteColumnGroupName + '" DisplayName="Page Content" Description="Page Content is a site column created by the Publishing feature. It is used on the Article Page Content Type as the content of the page." Type="HTML" Required="FALSE" Sealed="TRUE" RichText="TRUE" RichTextMode="FullHtml" />',

];
INTRANET.Schema.ContentTypes = [{
        Name: 'Latest_News1',
        ParentContenttype: '0x010100C568DB52D9D0A14D9B2FDCC96666E9F2007948130EC3DB064584E219954237AF39', //Document Content Type
        Description: 'Latest_News1',
        IsLookUpExist: false,
        FieldsInternalName: ['NewsTitle', 'NewsPublishedDate', 'NewsAuthor', 'NewsBanner', 'PublishingPageContent']
    },
    {
        Name: 'Global_Navigation',
        ParentContenttype: '0x01', //Announcement Content Type
        Description: 'Global Navigation',
        IsLookUpExist: false,
        FieldsInternalName: ['NavigationLevel', 'Sequence', 'IsActiveLink', 'OpenInNewTab', 'NavigationLink', 'Parent']
    },
    {
        Name: 'Intranet_Survey',
        ParentContenttype: '0x01', //Item Content Type
        Description: 'Survey List',
        IsLookUpExist: false,
        FieldsInternalName: ['SurveyDescription', 'SurveyLink', 'EndDate', 'ManageQuestions', 'ViewResponse', 'EmailNotification', 'IsPoll']
    },
    {
        Name: 'Latest_Update',
        ParentContenttype: '0x0104', //Announcement Content Type
        Description: 'Latest Updates',
        IsLookUpExist: false,
        FieldsInternalName: ['BannerImage', 'HomeThumbnail', 'ThumbnailImage', 'ContentAuthor', 'EmailNotification', 'StoryCopy']
    },
    {
        Name: 'Organizational_Announcement',
        ParentContenttype: '0x0104', //Announcement Content Type
        Description: 'Organizational_Announcement',
        IsLookUpExist: false,
        FieldsInternalName: ['BannerImage', 'ContentAuthor', 'ContentPriority', 'EmailNotification', 'AnnouncementDetails']
    },
    {
        Name: 'Mark_Your_Calendar',
        ParentContenttype: '0x0102', //Event Content Type
        Description: 'Mark_Your_Calendar',
        IsLookUpExist: false,
        FieldsInternalName: ['EmailNotification', 'EmployeeName']
    },
    {
        Name: 'News_Bulletin',
        ParentContenttype: '0x0104', //Announcement Content Type
        Description: 'News_Bulletin',
        IsLookUpExist: false,
        FieldsInternalName: ['ContentPriority', 'EmailNotification']
    },
    {
        Name: 'Voice_of_Management',
        ParentContenttype: '0x01', //Item Content Type
        Description: 'Voice_of_Management',
        IsLookUpExist: false,
        FieldsInternalName: ['ContentAuthor', 'ContentPriority', 'EmailNotification', 'ManagementMessage']
    },
    {
        Name: 'BirthDay',
        ParentContenttype: '0x0102', //Event Content Type
        Description: 'BirthDay',
        IsLookUpExist: false,
        FieldsInternalName: ['EmailNotification', 'EmployeeName']
    },
    {
        Name: 'Anniversary',
        ParentContenttype: '0x0102', //Event Content Type
        Description: 'Anniversary',
        IsLookUpExist: false,
        FieldsInternalName: ['EmailNotification', 'EmployeeName']
    },
    {
        Name: 'New_Joinee',
        ParentContenttype: '0x0102', //Event Content Type
        Description: 'New_Joinee',
        IsLookUpExist: false,
        FieldsInternalName: ['EmailNotification', 'EmployeeName']
    },
    {
        Name: 'Farewell',
        ParentContenttype: '0x0102', //Event Content Type
        Description: 'Farewell',
        IsLookUpExist: false,
        FieldsInternalName: ['EmailNotification', 'EmployeeName']
    },
    {
        Name: 'Issue Tracking',
        ParentContenttype: '0x01', //Event Content Type
        Description: 'Issue Tracking',
        IsLookUpExist: false,
        FieldsInternalName: ['AuditType', 'AuditCategory', 'AuditTitle', 'AuditAssignmentNoticeNo', 'AuditReportDate', 'AuditIssueSummary',
            'IssueSignificance', 'IssueDetails', 'AuditRecommendation', 'AgreedCorrectiveAction', 'OriginalDueDate',
            'RevisedDueDate', 'PersonToAddress', 'IssueStatus', 'DueStatus', 'IssueResolutionComments', 'IssueClosureDate',
            'SupervisoryLevel', 'Auditors', 'AuditManagers', 'AuditorRemarks', 'IssueOwner', 'ParentIssueID', 'ParentAuditTitle'
        ]
    },

    {
        Name: 'Application_Configuration',
        ParentContenttype: '0x01', //Item Content Type
        Description: 'Application Configurations',
        IsLookUpExist: false,
        FieldsInternalName: ['Key', 'Value', 'Category']
    }
];
INTRANET.Schema.Lists = [{
        Name: 'Application_Configuration',
        Type: 'genericList',
        TemplateID: '100',
        ContentType: 'Application_Configuration',
        ContentTypeID: '', //This parameter value getting insert dynamically while app execution
        SetAsDefaultContentType: true,
        IsLookUpExist: false,
        IsUniquePermission: false,
        Fields: [],
        Items: [
            [{
                    Key: "Category",
                    Value: "Navigation"
                }, {
                    Key: "Key",
                    Value: "Navigation_Term_Set_ID"
                },
                {
                    Key: "Value",
                    Value: "566255d4-9d35-41e1-9b50-d3558136cdd1"
                }
            ],
            [{
                    Key: "Category",
                    Value: "Navigation"
                }, {
                    Key: "Key",
                    Value: "NAVIGATION_FROM_LIST"
                },
                {
                    Key: "Value",
                    Value: "YES"
                }
            ]
        ]
    },
    {
        Name: 'IntranetSurvey',
        Type: 'genericList',
        TemplateID: '100',
        ContentType: 'Intranet_Survey',
        ContentTypeID: '', //This parameter value getting insert dynamically while app execution
        SetAsDefaultContentType: true,
        IsLookUpExist: false,
        IsUniquePermission: false,
        Fields: [],
        Items: []
    },
    {
        Name: 'LatestNews1',
        Type: 'documentLibrary',
        TemplateID: '101',
        ContentType: 'Latest_News1',
        ContentTypeID: '', //This parameter value getting insert dynamically while app execution
        SetAsDefaultContentType: true,
        IsLookUpExist: false,
        IsUniquePermission: false,
        Fields: [],
        Items: []
    },
    {
        Name: 'GlobalNavigation',
        Type: 'genericList',
        TemplateID: '101',
        ContentType: 'Global_Navigation',
        ContentTypeID: '', //This parameter value getting insert dynamically while app execution
        SetAsDefaultContentType: true,
        IsLookUpExist: false,
        IsUniquePermission: false,
        Fields: [],
        Items: [
            [{
                    Key: 'Title',
                    Value: 'Google'
                }, {
                    Key: 'NavigationLevel',
                    Value: 'Level 1'
                }, {
                    Key: 'Sequence',
                    Value: '1'
                },
                {
                    Key: 'IsActiveLink',
                    Value: 'Yes'
                },
                {
                    Key: 'OpenInNewTab',
                    Value: 'Yes'
                }, {
                    Key: 'NavigationLink',
                    Value: 'www.google.com'
                }
            ],
            [{
                    Key: 'Title',
                    Value: 'Yahoo'
                }, {
                    Key: 'NavigationLevel',
                    Value: 'Level 1'
                }, {
                    Key: 'Sequence',
                    Value: '2'
                },
                {
                    Key: 'IsActiveLink',
                    Value: 'Yes'
                },
                {
                    Key: 'OpenInNewTab',
                    Value: 'Yes'
                }, {
                    Key: 'NavigationLink',
                    Value: 'www.Yahoo.com'
                }
            ],
        ]
    }, {
        Name: 'LatestUpdates',
        Type: 'genericList',
        TemplateID: '101',
        ContentType: 'Latest_Update',
        ContentTypeID: '', //This parameter value getting insert dynamically while app execution
        SetAsDefaultContentType: true,
        IsLookUpExist: false,
        IsUniquePermission: false,
        Fields: [],
        Items: []
    }, {
        Name: 'OrganizationalAnnouncements',
        Type: 'genericList',
        TemplateID: '101',
        ContentType: 'Organizational_Announcement',
        ContentTypeID: '', //This parameter value getting insert dynamically while app execution
        SetAsDefaultContentType: true,
        IsLookUpExist: false,
        IsUniquePermission: false,
        Fields: [],
        Items: []
    }, {
        Name: 'Mark Your Calendar',
        Type: 'events',
        TemplateID: '106',
        ContentType: 'Mark_Your_Calendar',
        ContentTypeID: '', //This parameter value getting insert dynamically while app execution
        SetAsDefaultContentType: true,
        IsLookUpExist: false,
        IsUniquePermission: false,
        Fields: [],
        Items: []
    }, {
        Name: 'NBBBulletins',
        Type: 'genericList',
        TemplateID: '100',
        ContentType: 'News_Bulletin',
        ContentTypeID: '', //This parameter value getting insert dynamically while app execution
        SetAsDefaultContentType: false,
        IsLookUpExist: false,
        IsUniquePermission: false,
        Fields: [],
        Items: []
    }, {
        Name: 'VoiceofManagement',
        Type: 'genericList',
        TemplateID: '100',
        ContentType: 'Voice_of_Management',
        ContentTypeID: '', //This parameter value getting insert dynamically while app execution
        SetAsDefaultContentType: true,
        IsLookUpExist: false,
        IsUniquePermission: false,
        Fields: [],
        Items: []
    }, {
        Name: 'NBBFamilyUpdates',
        Type: 'genericList',
        TemplateID: '100',
        ContentType: 'BirthDay',
        ContentTypeID: '', //This parameter value getting insert dynamically while app execution
        SetAsDefaultContentType: false,
        IsLookUpExist: false,
        IsUniquePermission: false,
        Fields: [],
        Items: []
    }, {
        Name: 'Quick Links',
        Type: 'links',
        TemplateID: '103',
        ContentType: 'Links',
        ContentTypeID: '', //This parameter value getting insert dynamically while app execution
        SetAsDefaultContentType: false,
        IsLookUpExist: false,
        IsUniquePermission: false,
        Fields: [],
        Items: []
    }, {
        Name: 'Issue Tracking',
        Type: 'genericList',
        TemplateID: '100',
        ContentType: 'Issue Tracking',
        ContentTypeID: '', //This parameter value getting insert dynamically while app execution
        SetAsDefaultContentType: false,
        IsLookUpExist: false,
        IsUniquePermission: false,
        Fields: [],
        Items: []
    }

];

INTRANET.Schema.AllCSS = [
    "bootstrap.min.css", "custom_styles.css", "font-awesome.min.css", "mega_menu.css",
    "NBB_Main.css", "NBB_Responsive.css"
];

INTRANET.Schema.AllJS = [
    "bootstrap.min.js", "jquery-3.1.0.min.js", "knockout-3.2.0.js", "moment.js", "pnp.min.js", "services.js"
];

INTRANET.Schema.AllHtmlFiles = [
    "carousel.html", "LatestUpdatesJSLink.html", "OrgAnnouncementJSLink.html", "Poll.html"
];

INTRANET.Schema.AllFontFiles = [
    "calibri.eot", "calibri.otf", "calibri.ttf", "calibri.woff", "FontAwesome.otf", "fontawesome-webfont.eot",
    "fontawesome-webfont.svg", "fontawesome-webfont.ttf", "fontawesome-webfont.woff", "fontawesome-webfont.woff2",
    "Gill Sans MT Bold.eot", "Gill Sans MT Bold.svg", "Gill Sans MT Bold.woff", "Gill Sans MT Bold.woff2",
    "Gill Sans MT.eot", "Gill Sans MT.svg", "Gill Sans MT.woff2", "glyphicons-halflings-regular.eot",
    "glyphicons-halflings-regular.svg", "glyphicons-halflings-regular.ttf", "glyphicons-halflings-regular.woff",
    "glyphicons-halflings-regular.woff2"
];
INTRANET.Schema.AllImagesFiles = [
    "NBB-intranet-logo.jpg", "NBB-logo.jpg"
];

INTRANET.Schema.AllTemplateFiles = [
    "Control_KPCUNews", "Control_LatestUpdates", "Control_VoiceOfMgtPaging", "Item_KPCU_Happenings",
    "Item_LatestUpdateHomePage", "Item_News", "Item_VoiceOfManagement", "control_Happenings"
];
INTRANET.Schema.AllPageLayouts = [
    "KPCU_NewsDetailsLayout", "DepartmentGenricPageLayout", "NBBHomePageLayout", "NBBLearningPageLayout", "NBBOneColPageLayout"
];

INTRANET.Schema.AllPages = [
    "AllArticles", "AllLatestNews", "Home", "IntranetSurvey", "LatestUpdateDetails", "LatestUpdates",
    "MarkYourCalendar", "OrgAnnouncementDetails", "OrgAnnouncements", "QuickLinks"
]