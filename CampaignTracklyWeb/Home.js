var app = angular.module('myApp', []);
app.controller('myCtrl', function ($scope) {




    function ShowLoader() {
        document.getElementById('ProgressBgDiv').style.display = 'block';
        document.getElementById('loader').style.display = 'block';
        if (!$scope.$$phase) {
            $scope.$apply();
        };
    };

    function HideLoader() {

        document.getElementById('ProgressBgDiv').style.display = 'none';
        document.getElementById('loader').style.display = 'none';
        if (!$scope.$$phase) {
            $scope.$apply();
        };
    };

    ShowLoader();


    var timeout;

    function LoadToast(msg, isEror) {
        $scope.isError = isEror;
        $scope.Message = msg;
        if (!$scope.$$phase) {
            $scope.$apply();
        }
        $(".toast").toast("show");
        timeout = setTimeout(HideToast, 3000);
    };

    function HideToast() {
        $(".toast").toast("hide");
    };



    /////////// Functino for code refresh everytime perfectly ///////////
    function createGuid() {
        return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
            var r = Math.random() * 16 | 0, v = c === 'x' ? r : (r & 0x3 | 0x8);
            return v.toString(16);
        });
    };

    var guid = createGuid();

    $scope.LoginDiv = true;
    $scope.MainPageDiv = true;
    $scope.NavBarDiv = true;
    $scope.StartedScreen = true;
    $scope.SelectedOption;
    var APIToken = null;
    var FirstTime;
    $scope.UsedSheetData = [];
    $scope.result_Links;


    var BaseURL = "https://devapp.campaigntrackly.com";
   // var BaseURL = "https://app.campaigntrackly.com";

    /////////// show the started screen to user ///////////
    var checkUser = window.localStorage.getItem("UserVisted");

    if (checkUser === null) {
        $scope.StartedScreen = false;
        $scope.LoginDiv = true;
        $scope.MainPageDiv = true;
        $scope.NavBarDiv = true;
        FirstTime = true;
    } else {
        $scope.StartedScreen = true;
    };

    $scope.StartAddin = function () {
        window.localStorage.setItem("UserVisted", "Visted");
        $scope.StartedScreen = true;
        $scope.LoginDiv = false;
    };


    Office.onReady(function () {

       

        $scope.cehckShet = function () {

            Excel.run(function (context) {
                let sheets = context.workbook.worksheets;
                sheets.load("items");
            
                let sheet = context.workbook.worksheets.getItem("Sheet1");
                sheet.load("name, position");
                sheet.activate();
                return context.sync().then(function () {

                   
                }).catch(function (error) {
                    console.log(error);

                });

            });
        };



        /////////// check user is logined or not ///////////
        var getFromLocal = window.localStorage.getItem("APIToken");
        if (getFromLocal != null) {
            getFromLocal = JSON.parse(getFromLocal);
            APIToken = getFromLocal.token;
        };

        /////////// check token expiration ///////////
        function isTokenExpired(token) {
            const base64Url = token.split(".")[1];
            const base64 = base64Url.replace(/-/g, "+").replace(/_/g, "/");
            const jsonPayload = decodeURIComponent(
                atob(base64)
                    .split("")
                    .map(function (c) {
                        return "%" + ("00" + c.charCodeAt(0).toString(16)).slice(-2);
                    })
                    .join("")
            );

            const { exp } = JSON.parse(jsonPayload);
            var expnew = exp * 1000;

            var ee = new Date(Date.now());
            var ef = new Date(expnew);
            if (ee > ef) {
                expired = true;
            }
            else {
                expired = false;
            };

            return expired
        };

        var tokenFreshed;

        //////////////////////// Refresh token ////////////////////////

        function RefreshToken(refshToken) {

            var settings = {
                "url": BaseURL +"/wp-json/campaigntrackly/v1/refresh_token",
                "method": "POST",
                "timeout": 0,
                "async": false,
                "headers": {
                    "Accept": "application/json",
                    "Content-Type": "application/x-www-form-urlencoded"
                },
                "data": {
                    "refresh_token": refshToken
                }
            };


            $.ajax(settings).done(function (response) {
                //  console.log(response);

                if (response.statusCode === 200) {
                    APIToken = response.data.token;
                    window.localStorage.setItem("APIToken", JSON.stringify(response.data));
                    HideLoader();
                    tokenFreshed = true;
                } else {
                    $scope.LoginDiv = false;
                    $scope.MainPageDiv = true;
                    $scope.NavBarDiv = true;
                    LoadToast(response.message)
                    tokenFreshed = false;

                    // HideLoader();
                };
                HideLoader();

            }).fail(function (error) {
                console.log(error);

                HideLoader();
                LoadToast("Connection Issue. Please contact support@campaigntrackly.com");

            });


        };



        //////////////////////// Sign In ////////////////////////
        $scope.SignIn = function () {
            ShowLoader();


            var settings = {
                "url": BaseURL +"/wp-json/campaigntrackly/v1/login",
                "method": "POST",
                "timeout": 0,
                "headers": {
                    "Content-Type": "application/x-www-form-urlencoded",
                },
                "data": {
                    "username": $scope.UserName.trim(),
                    "password": $scope.UserPassword.trim()
                }
            };

            $.ajax(settings).done(function (response) {
                //  console.log(response);

                if (response.statusCode === 200) {

                    $scope.UserName = undefined;
                    $scope.UserPassword = undefined;

                    $scope.LoginDiv = true;
                    $scope.MainPageDiv = false;
                    $scope.NavBarDiv = false;
                    if (response.data.token) {
                        APIToken = response.data.token;
                        window.localStorage.setItem("APIToken", JSON.stringify(response.data));
                        $scope.getTagTemplates();

                    };


                } else {
                    $scope.LoginDiv = false;
                    $scope.MainPageDiv = true;
                    $scope.NavBarDiv = true;
                    // HideLoader();
                };



            }).fail(function (error) {
                console.log(error);
                if (error.responseJSON.statusCode) {

                    if (error.responseJSON.statusCode === 403 || error.responseJSON.code === "application_passwords_disabled") {
                        LoadToast(error.responseJSON.message, true);
                    } else {
                        LoadToast("Login Failed", true);
                    };
                };

                HideLoader();

            });

        };




        ////////////////// All Column Autofill //////////////////
        function AllSheetAutoFill() {
            Excel.run(function (context) {

                let myWorkbook = context.workbook;
                let sheet = myWorkbook.worksheets.getActiveWorksheet();

                let range = sheet.getUsedRange();
                range.format.autofitColumns();


                return context.sync().then(function () {
                    // console.log("Autofill");
                });
            });
        };

        AllSheetAutoFill();


        //////////////////////// Get tag_templates for dropdown ////////////////////////
        $scope.getTagTemplates = function () {
            ShowLoader();

            $.ajax({
                url: BaseURL +"/wp-json/campaigntrackly/v1/tag_templates",
                method: "GET",
                headers: {
                    "accept": "application/json",
                    "Authorization": "Bearer " + APIToken
                },
                success: function (response) {
                    //  console.log(response);

                    $scope.Tag_TemplatesArr = response;
                    $scope.SelectedOption = "Dummy";

                    HideLoader();
                    if (!$scope.$$phase) {
                        $scope.$apply();
                    };
                },
                error: function (error) {
                    console.log(error);
                    HideLoader();

                    if (error.responseJSON) {
                        if (error.responseJSON.statusCode === 403 && error.responseJSON.message === "Expired token") {
                            RefreshToken(getFromLocal.refresh_token);
                            if (tokenFreshed) {
                                $scope.getTagTemplates();
                            };
                        } else {
                            LoadToast("Connection Issue. Please contact support@campaigntrackly.com");
                        }
                    } else {
                        LoadToast("Connection Issue. Please contact support@campaigntrackly.com");
                    }

                }
            });

        };



        function alphaOnly(a) {
            var b = '';
            for (var i = 0; i < a.length; i++) {
                if (a[i] >= 'A' && a[i] <= 'z') b += a[i];
            }
            return b;
        };

        function nextLetter(s) {
            return s.replace(/([a-zA-Z])[^a-zA-Z]*$/, function (a) {
                var c = a.charCodeAt(0);
                switch (c) {
                    case 90: return 'A';
                    case 122: return 'a';
                    default: return String.fromCharCode(++c);
                }
            });
        };



        ///////////////////////////////////// Apply Template////////////////////////////////////////////


        var OtherTags = [];
        var indxOfCampName;
        var indxOfURL;
        var indxOfContentTag;
        var indxOfSource;
        var indxOfMedium;
        var indxOfTerms;
        var Scenario;
        var AllNameUrlArr = [];
        var CamNameURLObj = {};
        var PrepareFinalArr = [];
        var PrepareDataApplyTemplate = {};
        var FinalSheetSet = [];
        var headerList = [];
        var CustomTagAPI = [];
        var SelctedTemTag = [];
        var ActiveSheet ;



        function GetAllCustTags() {

            //////////////////////// Get Custom Tags ////////////////////////

            //$.ajax({
            //    url: BaseURL + "/wp-json/campaigntrackly/v1/custom_tags_names",
            //    method: "GET",
            //    async: false,
            //    headers: {
            //        "accept": "application/json",
            //        "Authorization": "Bearer " + APIToken
            //    },
            //    success: function (response) {
            //        console.log(response);

            //        for (let i = 0; i < response.length; i++) {
            //            CustomTagAPI.push(response[i].custom.toLowerCase());
            //        };

                
            //    },
            //    error: function (error) {
            //        console.log(error);
            //    }
            //});




            $.ajax({
                url: BaseURL + "/wp-json/campaigntrackly/v1/tag_templates",
                method: "GET",
                async: false,
                headers: {
                    "accept": "application/json",
                    "Authorization": "Bearer " + APIToken
                },
                success: function (response) {
                    //  console.log(response);
                    for (let i = 0; i < response.length; i++) {
                        //console.log(response[i].id);
                        if (response[i].id === $scope.SelectedOption.id) {
                            SelctedTemTag = response[i].custom;
                        };
                    };

                    for (var m = 0; m < SelctedTemTag.length; m++) {
                        CustomTagAPI.push(SelctedTemTag[m].title.toLowerCase());
                    };

                    //  console.log(CustomTagAPI);
                },
                error: function (error) {
                    console.log(error);
                    HideLoader();


                    if (error.status != 200) {

                        if (error.responseJSON.statusCode === 403 && error.responseJSON.message === "Expired token") {
                            RefreshToken(getFromLocal.refresh_token);
                            GetAllCustTags();
                        } else {
                            LoadToast("Connection Issue. Please contact support@campaigntrackly.com");
                        } 
                     
                    } else {
                        LoadToast("Connection Issue. Please contact support@campaigntrackly.com");
                    }

                }
            });


        };


      //  console.log("Working File");
        $scope.ApplyTemplate = function () {

            ShowLoader();
            CustomTagAPI = [];
            SelctedTemTag = [];
            $scope.UsedSheetValues = [];

            Excel.run(function (context) {
                let sheetActCall = context.workbook.worksheets.getActiveWorksheet();
                sheetActCall.load("name");

                return context.sync().then(function () {
             
                    if (sheetActCall.name.includes("Result_")) {
                   
                        HideLoader();
                        LoadToast("Connection Issue. Please contact support@campaigntrackly.com");

                    } else {

                        AllNameUrlArr = [];
                        OnlyNameArr = [];
                        CamNameURLObj = {};
                        PrepareFinalArr = [];
                        PrepareDataApplyTemplate = {};
                        FinalSheetSet = [];
                        AllTagData = [];
                        LinksOfSncdSca = [];
                        checkRes = false;

                        Excel.run(function (context) {

                            let myWorkbook = context.workbook;
                            let sheet = myWorkbook.worksheets.getActiveWorksheet();

                            let range = sheet.getUsedRange();

                            GetAllCustTags();

                            return context.sync().then(function () {
                                var DataResults = range.load("values");

                                return context.sync().then(function () {
                                    // console.log(DataResults.values);

                                    allData = DataResults.values;

                                    for (let m = 0; m < allData.length; m++) {

                                        const allEmptyOrNewline = allData[m].every(item => item === "" || item === "\n");

                                        if (!allEmptyOrNewline) {
                                            $scope.UsedSheetValues.push(allData[m]);
                                        } else {
                                          //  console.log("All items in the array are equal to empty strings or newline characters.");
                                        };
                                    };


                                    var lowerCaseHeadArr = $scope.UsedSheetValues[0];

                                  var  headerListLow = lowerCaseHeadArr.map(item => item.toLowerCase());

                                    function replaceMultipleSpaces(str) {
                                        return str.replace(/\s{2,}/g, ' ');
                                    };

                                    function replaceMultipleSpacesInArray(array) {
                                        return array.map(function (iteme) {
                                            return replaceMultipleSpaces(iteme);
                                        });
                                    };

                                    const headerList = replaceMultipleSpacesInArray(headerListLow);                                   
                             

                                    //////////////////////// Check Scenario ////////////////////////

                                    if (headerList.includes("campaign name") && headerList.includes("url") && !headerList.includes('') && !headerList.includes("content") && !headerList.includes("terms") && !headerList.includes("source") && !headerList.includes("medium")) {

                                        Scenario = "First Scenario";

                                        for (let i = 0; i < headerList.length; i++) {

                                            if (headerList[i] === "campaign name") {
                                                indxOfCampName = i;
                                            };
                                            if (headerList[i] === "url") {
                                                indxOfURL = i;
                                            };
                                        };

                                    }
                                    else {
                                        Scenario = "Secound Scenario";

                                        OtherTags = [];


                                        var checkCountCampName = [];
                                        var objToCamName = {};
                                        const itemToCheck = "campaign name";

                                        for (let m = 0; m < headerList.length; m++) {
                                            if (headerList[m] === itemToCheck) {
                                                objToCamName = {
                                                    "headName": headerList[m],
                                                    "CampIndx": m
                                                };
                                                checkCountCampName.push(objToCamName);
                                                objToCamName = {};
                                            };
                                        };


                                        for (let i = 0; i < headerList.length; i++) {
                                            if (headerList[i] === "campaign name") {
                                                if (checkCountCampName.length === 1) {
                                                    indxOfCampName = i;
                                                };
                                                if (i === 0) {
                                                    indxOfCampName = i;
                                                };
                                            } else if (headerList[i] === "url") {
                                                indxOfURL = i;
                                            } else if (headerList[i] === "content") {
                                                indxOfContentTag = i;
                                            } else if (headerList[i] === "medium") {
                                                indxOfMedium = i;
                                            } else if (headerList[i] === "terms") {
                                                indxOfTerms = i;
                                            } else if (headerList[i] === "source") {
                                                indxOfSource = i;
                                            } else {
                                                // if (headerList[i] != "result" && headerList[i] != "short links" && headerList[i] != "date") {
                                                if (CustomTagAPI.includes(headerList[i])) {
                                                    var CustomTagObj = {
                                                        "TagName": headerList[i],
                                                        "TagIndex": i
                                                    };
                                                    OtherTags.push(CustomTagObj);
                                                    CustomTagObj = {};
                                                };

                                            };
                                        };
                                      //  console.log(OtherTags);
                                    

                                    };

                                    //////////////////////// First Scenario ////////////////////////

                                    if (Scenario === "First Scenario") {

                                        for (var n = 1; n < $scope.UsedSheetValues.length; n++) {
                                            if ($scope.UsedSheetValues[n][indxOfCampName] != "" || $scope.UsedSheetValues[n][indxOfURL] != "") {
                                                CamNameURLObj = {
                                                    "CampaignName": $scope.UsedSheetValues[n][indxOfCampName],
                                                    "CampaignURL": $scope.UsedSheetValues[n][indxOfURL]
                                                };
                                                AllNameUrlArr.push(CamNameURLObj);
                                                CamNameURLObj = {};
                                            };
                                        };

                                        for (let i = 0; i < AllNameUrlArr.length; i++) {
                                            PrepareDataApplyTemplate = {
                                                "template_id": $scope.SelectedOption.id,
                                                "campaign_name": AllNameUrlArr[i].CampaignName,
                                                "links": [
                                                    AllNameUrlArr[i].CampaignURL
                                                ]
                                            };
                                            PrepareFinalArr.push(PrepareDataApplyTemplate);
                                            PrepareDataApplyTemplate = {};
                                        };


                                        $.ajax({
                                            url: BaseURL + "/wp-json/campaigntrackly/v1/apply_template",
                                            method: "POST",
                                            headers: {
                                                "accept": "application/json",
                                                "Authorization": "Bearer " + APIToken
                                            },
                                            data: JSON.stringify(PrepareFinalArr),
                                            success: function (response) {
                                                //   console.log(response);


                                                if (response.code) {
                                                    if (response.code === "401") {
                                                        HideLoader();
                                                        LoadToast(response.response);

                                                    };
                                                };

                                                if (response.code != "401") {

                                                    $scope.result_Links = response;

                                                    if ($scope.result_Links[0].links.length > 0) {

                                                        FinalSheetSet = [];
                                                        var UrlItem = [];

                                                        //for (let i = 0; i < $scope.result_Links.length; i++) {
                                                        //    for (let m = 0; m < $scope.result_Links[i].links.length; m++) {
                                                        //        var ForSheetSet = [AllNameUrlArr[i].CampaignName, AllNameUrlArr[i].CampaignURL, $scope.result_Links[i].links[m], $scope.result_Links[i].short_links[m], $scope.result_Links[i].date];
                                                        //        FinalSheetSet.push(ForSheetSet);

                                                        //    };
                                                        //};


                                                        for (var i = 0; i < $scope.UsedSheetValues.length;) {
                                                            if (i != 0) {
                                                                for (var m = 0; m < $scope.result_Links.length; m++) {
                                                                    if ($scope.result_Links[m].links.length > 0) {
                                                                        for (var n = 0; n < $scope.result_Links[m].links.length; n++) {
                                                                            FinalSheetSet.push($scope.UsedSheetValues[i]);
                                                                        };
                                                                        i++;
                                                                    } else {
                                                                        FinalSheetSet.push($scope.UsedSheetValues[i]);
                                                                    };
                                                                };
                                                            } else {
                                                                FinalSheetSet.push($scope.UsedSheetValues[i]);
                                                                i++
                                                            };
                                                        };

                                                        for (var m = 0; m < $scope.result_Links.length; m++) {
                                                            if ($scope.result_Links[m].links.length > 0) {
                                                                for (var n = 0; n < $scope.result_Links[m].links.length; n++) {
                                                                    UrlItem.push([$scope.result_Links[m].links[n], $scope.result_Links[m].short_links[n], $scope.result_Links[m].date])
                                                                };
                                                            } else {
                                                                UrlItem.push(['', '', $scope.result_Links[m].date]);
                                                            };
                                                        };

                                                        var lastColName = "";
                                                        HeadNames = $scope.UsedSheetValues[0];
                                                        var markers = [];

                                                        for (var n = 0; n < HeadNames.length; n++) {
                                                            var Aplhabet = (n + 10).toString(36).toUpperCase();
                                                            markers[i] = sheet.getRange(Aplhabet + 1);
                                                            markers[i].values = HeadNames[n];
                                                            if (n < HeadNames.length) {
                                                                if (HeadNames[n] != "Result" && HeadNames[n] != "Short Links" && HeadNames[n] != "Date") {
                                                                    lastColName = Aplhabet;
                                                                };
                                                            };
                                                        };



                                                    


                                                        Excel.run(function (context) {
                                                            let Actsheet = context.workbook.worksheets.getActiveWorksheet();
                                                            Actsheet.load("name");

                                                            let sheets = context.workbook.worksheets;
                                                            sheets.load("items/name");

                                                            return context.sync().then(function () {

                                                                var checkRes;
                                                                for (var i = 0; i < sheets.items.length; i++) {
                                                                    ActiveSheet = Actsheet.name;
                                                                    var activeSheetRes = "Result_" + ActiveSheet;
                                                                    if (sheets.items[i].name === activeSheetRes) {
                                                                        checkRes = true;
                                                                        break;
                                                                    } else {
                                                                        checkRes = false;
                                                                    };
                                                                }



                                                                if (checkRes === true) {

                                                                    let ResultSheet = context.workbook.worksheets.getItem("Result_" + ActiveSheet);

                                                                    var UsdRangeRes = ResultSheet.getUsedRange();
                                                                    UsdRangeRes.clear();

                                                                    return context.sync().then(function () {


                                                                        var NextColumnForResult = nextLetter(lastColName);
                                                                        var NextColumnForShort = nextLetter(NextColumnForResult);
                                                                        var NextColumnForDate = nextLetter(NextColumnForShort);
                                                                        var rangeForResHead = ResultSheet.getRange(NextColumnForResult + 1 + ":" + NextColumnForDate + 1);
                                                                        rangeForResHead.values = [["Result", "Short Links", "Date"]];
                                                                        var toRangeLink = UrlItem.length + 1;
                                                                        var range_Link = NextColumnForResult + 2 + ":" + NextColumnForDate + toRangeLink;
                                                                        var rangeForResLink = ResultSheet.getRange(range_Link);


                                                                        let data = FinalSheetSet;
                                                                        var FROM = 1;
                                                                        var TO = FROM + data.length - 1;
                                                                        var RANEG = "A" + FROM.toString() + ":" + Aplhabet + TO.toString();
                                                                        let range = ResultSheet.getRange(RANEG);
                                                                        range.formulas = data;
                                                                        range.format.autofitColumns();








                                                                        //var ColmnCamName = "A";
                                                                        //var ColmnCamURL = nextLetter(ColmnCamName);
                                                                        //var ColmnRes = nextLetter(ColmnCamURL);
                                                                        //var ColmnShortLink = nextLetter(ColmnRes);
                                                                        //var ColmnDate = nextLetter(ColmnShortLink);

                                                                        //var rangeForCampainNameHead = ResultSheet.getRange(ColmnCamName + 1);
                                                                        //rangeForCampainNameHead.values = "Campaign Name"
                                                                        //var rangeForCampainURLHead = ResultSheet.getRange(ColmnCamURL + 1);
                                                                        //rangeForCampainURLHead.values = "URL"
                                                                        //var rangeForCampainResHead = ResultSheet.getRange(ColmnRes + 1);
                                                                        //if (headerList.includes("result")) {
                                                                        //    var ResLengthTo = AllNameUrlArr.length + 1;
                                                                        //    var ValuesOfRes = ResultSheet.getRange(ColmnRes + 1 + ":" + ColmnRes + ResLengthTo);
                                                                        //    ValuesOfRes.clear();
                                                                        //    rangeForCampainResHead.values = "Result";
                                                                        //} else {
                                                                        //    rangeForCampainResHead.values = "Result";
                                                                        //};


                                                                        //var rangeForShortLinkHead = ResultSheet.getRange(ColmnShortLink + 1);

                                                                        //if (headerList.includes("short links")) {
                                                                        //    var ResLengthTo = AllNameUrlArr.length + 1;
                                                                        //    var ValuesOfShortLink = ResultSheet.getRange(ColmnShortLink + 1 + ":" + ColmnShortLink + ResLengthTo);
                                                                        //    ValuesOfShortLink.clear();
                                                                        //    rangeForShortLinkHead.values = "Short Links";
                                                                        //} else {
                                                                        //    rangeForShortLinkHead.values = "Short Links";
                                                                        //};


                                                                        //var rangeForDateHead = ResultSheet.getRange(ColmnDate + 1);

                                                                        //if (headerList.includes("date")) {
                                                                        //    var ResLengthTo = AllNameUrlArr.length + 1;
                                                                        //    var ValuesOfDate = ResultSheet.getRange(ColmnDate + 1 + ":" + ColmnDate + ResLengthTo);
                                                                        //    ValuesOfDate.clear();
                                                                        //    rangeForDateHead.values = "Date";
                                                                        //} else {
                                                                        //    rangeForDateHead.values = "Date";
                                                                        //};


                                                                        //let data = FinalSheetSet;
                                                                        //var FROM = 2;
                                                                        //var TO = FROM + data.length - 1;
                                                                        //var RANEG = ColmnCamName + FROM.toString() + ":" + ColmnDate + TO.toString();
                                                                        //let range = ResultSheet.getRange(RANEG);
                                                                        //range.formulas = data;
                                                                        //range.format.autofitColumns();
                                                                        //var range_Val_Links = ColmnRes + FROM + ":" + ColmnRes + TO.toString();
                                                                        //var ValOfResLinks = ResultSheet.getRange(range_Val_Links);

                                                                        var range_LinksRes = NextColumnForResult + 2 + ":" + NextColumnForResult + toRangeLink;
                                                                        var rangeValOfLinks = ResultSheet.getRange(range_LinksRes);

                                                                        rangeValOfLinks.format.wrapText = true;
                                                                        rangeValOfLinks.format.columnWidth = 250;



                                                                        //   let sheet = context.workbook.worksheets.getItem("Sheet1");
                                                                        //   sheet.load("name, position");
                                                                        ResultSheet.activate();

                                                                        return context.sync().then(function () {
                                                                            rangeForResLink.values = UrlItem;
                                                                            rangeForResLink.format.autofitColumns();
                                                                            HideLoader();

                                                                        });


                                                                    });


                                                                } else {
                                                                    Excel.run(function (context) {

                                                                        let sheets = context.workbook.worksheets;

                                                                        let sheet = sheets.add("Result_" + ActiveSheet);
                                                                        sheet.load("name, position");

                                                                        return context.sync().then(function () {

                                                                            let ResultSheet = context.workbook.worksheets.getItem("Result_" + ActiveSheet);


                                                                            var ColmnCamName = "A";
                                                                            var ColmnCamURL = nextLetter(ColmnCamName);
                                                                            var ColmnRes = nextLetter(ColmnCamURL);
                                                                            var ColmnShortLink = nextLetter(ColmnRes);
                                                                            var ColmnDate = nextLetter(ColmnShortLink);

                                                                            var rangeForCampainNameHead = ResultSheet.getRange(ColmnCamName + 1);
                                                                            rangeForCampainNameHead.values = "Campaign Name"
                                                                            var rangeForCampainURLHead = ResultSheet.getRange(ColmnCamURL + 1);
                                                                            rangeForCampainURLHead.values = "URL"
                                                                            var rangeForCampainResHead = ResultSheet.getRange(ColmnRes + 1);
                                                                            if (headerList.includes("result")) {
                                                                                var ResLengthTo = AllNameUrlArr.length + 1;
                                                                                var ValuesOfRes = ResultSheet.getRange(ColmnRes + 1 + ":" + ColmnRes + ResLengthTo);
                                                                                ValuesOfRes.clear();
                                                                                rangeForCampainResHead.values = "Result";
                                                                            } else {
                                                                                rangeForCampainResHead.values = "Result";
                                                                            };


                                                                            var rangeForShortLinkHead = ResultSheet.getRange(ColmnShortLink + 1);

                                                                            if (headerList.includes("short links")) {
                                                                                var ResLengthTo = AllNameUrlArr.length + 1;
                                                                                var ValuesOfShortLink = ResultSheet.getRange(ColmnShortLink + 1 + ":" + ColmnShortLink + ResLengthTo);
                                                                                ValuesOfShortLink.clear();
                                                                                rangeForShortLinkHead.values = "Short Links";
                                                                            } else {
                                                                                rangeForShortLinkHead.values = "Short Links";
                                                                            };


                                                                            var rangeForDateHead = ResultSheet.getRange(ColmnDate + 1);

                                                                            if (headerList.includes("date")) {
                                                                                var ResLengthTo = AllNameUrlArr.length + 1;
                                                                                var ValuesOfDate = ResultSheet.getRange(ColmnDate + 1 + ":" + ColmnDate + ResLengthTo);
                                                                                ValuesOfDate.clear();
                                                                                rangeForDateHead.values = "Date";
                                                                            } else {
                                                                                rangeForDateHead.values = "Date";
                                                                            };




                                                                            let data = FinalSheetSet;
                                                                            var FROM = 2;
                                                                            var TO = FROM + data.length - 1;
                                                                            var RANEG = ColmnCamName + FROM.toString() + ":" + ColmnDate + TO.toString();
                                                                            let range = ResultSheet.getRange(RANEG);
                                                                            range.formulas = data;
                                                                            range.format.autofitColumns();
                                                                            var range_Val_Links = ColmnRes + FROM + ":" + ColmnRes + TO.toString();
                                                                            var ValOfResLinks = ResultSheet.getRange(range_Val_Links);

                                                                            ValOfResLinks.format.columnWidth = 250;
                                                                            ValOfResLinks.format.wrapText = true;

                                                                            ResultSheet.activate();

                                                                            return context.sync().then(function () {
                                                                                HideLoader();

                                                                            });

                                                                        });
                                                                    });
                                                                };


                                                            }).catch(function (error) {
                                                                console.log(error);

                                                            });

                                                        });


                                                    } else {
                                                        HideLoader();
                                                        LoadToast("Connection Issue. Please contact support@campaigntrackly.com");
                                                    };


                                                } else {
                                                    HideLoader();
                                                    LoadToast(response.response);
                                                };

                                                if (!$scope.$$phase) {
                                                    $scope.$apply();
                                                };
                                            },
                                            error: function (error) {
                                                if (error.status != 200) {
                                                    if (error.responseJSON.statusCode === 403 && error.responseJSON.message === "Expired token") {
                                                        RefreshToken(getFromLocal.refresh_token);
                                                        ShowLoader();
                                                        $scope.ApplyTemplate();
                                                    }
                                                    else {
                                                        LoadToast("Connection Issue. Please contact support@campaigntrackly.com");
                                                    };
                                                } else {
                                                    LoadToast("Connection Issue. Please contact support@campaigntrackly.com");

                                                }
                                                
                                                HideLoader();
                                            }
                                        });
                                    };


                                    //////////////////////// Second Scenario ////////////////////////

                                    if (Scenario === "Secound Scenario") {


                                        for (var n = 1; n < $scope.UsedSheetValues.length; n++) {

                                            if ($scope.UsedSheetValues[n][indxOfCampName] != "" || $scope.UsedSheetValues[n][indxOfURL] != "") {

                                                CamNameURLObj = {};
                                                CamNameURLObj["CampaignName"] = ($scope.UsedSheetValues[n][indxOfCampName] ? $scope.UsedSheetValues[n][indxOfCampName] : '');
                                                CamNameURLObj["CampaignURL"] = ($scope.UsedSheetValues[n][indxOfURL] ? $scope.UsedSheetValues[n][indxOfURL] : '');
                                                CamNameURLObj["ContentTag"] = ($scope.UsedSheetValues[n][indxOfContentTag] ? $scope.UsedSheetValues[n][indxOfContentTag] : '');
                                                CamNameURLObj["UtmMedium"] = ($scope.UsedSheetValues[n][indxOfMedium] ? $scope.UsedSheetValues[n][indxOfMedium] : '');
                                                CamNameURLObj["UtmTerm"] = ($scope.UsedSheetValues[n][indxOfTerms] ? $scope.UsedSheetValues[n][indxOfTerms] : '');
                                                CamNameURLObj["UtmSource"] = ($scope.UsedSheetValues[n][indxOfSource] ? $scope.UsedSheetValues[n][indxOfSource] : '');

                                                AllNameUrlArr.push(CamNameURLObj);
                                                CamNameURLObj = {};
                                            };
                                        };

                                        var OtherTagValArr = [];

                                        for (var l = 0; l < OtherTags.length; l++) {
                                            for (var i = 1; i < $scope.UsedSheetValues.length; i++) {
                                                var OtherTagVal = $scope.UsedSheetValues[i][OtherTags[l].TagIndex];
                                                var ObjOfOther = {};



                                                if (OtherTagValArr.length > 0) {
                                                    if (OtherTags[l].TagName != Object.keys(OtherTagValArr[OtherTagValArr.length - 1])) {
                                                        ObjOfOther[OtherTags[l].TagName] = [OtherTagVal]
                                                    } else {
                                                        var lastIndexTagName = Object.keys(OtherTagValArr[OtherTagValArr.length - 1]);
                                                        OtherTagValArr[OtherTagValArr.length - 1][lastIndexTagName[0]].push(OtherTagVal);
                                                        lastIndexTagName = [];
                                                        ObjOfOther = null;
                                                    };
                                                } else {
                                                    ObjOfOther[OtherTags[l].TagName] = [OtherTagVal]
                                                };

                                                if (ObjOfOther != null) {
                                                    OtherTagValArr.push(ObjOfOther);
                                                };
                                            };
                                        };


                                        var PreCustTagForSet = [];
                                        var CustTagForSet = [];

                                        for (let i = 0; i < OtherTagValArr.length; i++) {
                                            var keyOfObj = Object.keys(OtherTagValArr[i]);
                                            var ArrOfTagItem = OtherTagValArr[i][keyOfObj[0]];
                                            for (let m = 0; m < ArrOfTagItem.length; m++) {
                                                keyOfObj = Object.keys(OtherTagValArr[i]);
                                                PreCustTagForSet.push(OtherTagValArr[i][keyOfObj[0]][m]);
                                                keyOfObj = "";
                                            };
                                            CustTagForSet.push(PreCustTagForSet);
                                            PreCustTagForSet = [];

                                        };

                                        var custArr = [];


                                        for (let i = 0; i < AllNameUrlArr.length; i++) {


                                            for (let m = 0; m < CustTagForSet.length; m++) {
                                                var CusHeadName = [OtherTags[m].TagName];
                                                if (!CusHeadName[0].includes("date")) {
                                                    custArr.push({ [OtherTags[m].TagName]: [CustTagForSet[m][i]] });
                                                } else {
                                                    var ChangeFormate = CustTagForSet[m][i];
                                                    custArr.push({ [OtherTags[m].TagName]: [getJsDateFromExcel(ChangeFormate)] });
                                                };
                                            };


                                            PrepareDataApplyTemplate = {
                                                "template_id": $scope.SelectedOption.id,
                                                "campaign_name": AllNameUrlArr[i].CampaignName,
                                                "links": [{
                                                    "link": AllNameUrlArr[i].CampaignURL,
                                                    "channels": {
                                                        "source": AllNameUrlArr[i].UtmSource,
                                                        "medium": AllNameUrlArr[i].UtmMedium,
                                                        "terms":
                                                            (AllNameUrlArr[i].UtmTerm === "" ? [] : [AllNameUrlArr[i].UtmTerm])

                                                    },
                                                    "content": AllNameUrlArr[i].ContentTag,
                                                    "custom": custArr
                                                }]
                                            };




                                            custArr = [];
                                            PrepareFinalArr.push(PrepareDataApplyTemplate);
                                            PrepareDataApplyTemplate = {};
                                        };

                                        //console.log(PrepareFinalArr);

                                        var settings = {
                                            "url": BaseURL +"/wp-json/campaigntrackly/v1/apply_template_new_tags",
                                            "method": "POST",
                                            "timeout": 0,
                                            "headers": {
                                                "Accept": "application/json",
                                                "Content-Type": "application/json",
                                                "Authorization": "Bearer " + APIToken
                                            },
                                            "data": JSON.stringify(PrepareFinalArr),
                                        };

                                        $.ajax(settings).done(function (result) {
                                            // console.log(result);


                                            if (result.code) {
                                                if (result.code === "401") {
                                                    HideLoader();
                                                    LoadToast(result.response);

                                                };
                                            };



                                            $scope.result_Links = result;



                                            if (result.code != "401") {

                                                if ($scope.result_Links[0].links.length > 0) {

                                                    FinalSheetSet = [];

                                                    var UrlItem = [];
                                                    OnlyNameArr = [];

                                                    for (var i = 0; i < $scope.UsedSheetValues.length;) {
                                                        if (i != 0) {
                                                            for (var m = 0; m < $scope.result_Links.length; m++) {
                                                                if ($scope.result_Links[m].links.length > 0) {
                                                                    for (var n = 0; n < $scope.result_Links[m].links.length; n++) {
                                                                        FinalSheetSet.push($scope.UsedSheetValues[i]);

                                                                    };
                                                                    i++;
                                                                } else {
                                                                    FinalSheetSet.push($scope.UsedSheetValues[i]);

                                                                };
                                                            };
                                                        } else {
                                                            FinalSheetSet.push($scope.UsedSheetValues[i]);
                                                            i++
                                                        };
                                                    };

                                                    for (var m = 0; m < $scope.result_Links.length; m++) {
                                                        if ($scope.result_Links[m].links.length > 0) {
                                                            for (var n = 0; n < $scope.result_Links[m].links.length; n++) {
                                                                UrlItem.push([$scope.result_Links[m].links[n], $scope.result_Links[m].short_links[n], $scope.result_Links[m].date])
                                                            };
                                                        } else {
                                                            UrlItem.push(['', '', $scope.result_Links[m].date]);

                                                        };

                                                    };





                                                    Excel.run(function (context) {
                                                        let Actsheet = context.workbook.worksheets.getActiveWorksheet();
                                                        Actsheet.load("name");

                                                        let sheets = context.workbook.worksheets;
                                                        sheets.load("items/name");

                                                        return context.sync().then(function () {

                                                            var checkRes;
                                                            for (var i = 0; i < sheets.items.length; i++) {
                                                                ActiveSheet = Actsheet.name;
                                                                var activeSheetRes = "Result_" + ActiveSheet;
                                                                if (sheets.items[i].name === activeSheetRes) {
                                                                    checkRes = true;
                                                                    break;
                                                                } else {
                                                                    checkRes = false;
                                                                };
                                                            };

                                                            if (checkRes === true) {

                                                                let ResultSheet = context.workbook.worksheets.getItem("Result_" + ActiveSheet);

                                                                var UsdRangeRes = ResultSheet.getUsedRange();
                                                                UsdRangeRes.clear();


                                                                var HeadNames = $scope.UsedSheetValues[0];
                                                                var markers = [];
                                                                var lastColName;
                                                                for (var n = 0; n < HeadNames.length; n++) {
                                                                    var Aplhabet = (n + 10).toString(36).toUpperCase();
                                                                    markers[i] = sheet.getRange(Aplhabet + 1);
                                                                    markers[i].values = HeadNames[n];
                                                                    if (n < HeadNames.length) {
                                                                        if (HeadNames[n] != "Result" && HeadNames[n] != "Short Links" && HeadNames[n] != "Date") {
                                                                            lastColName = Aplhabet;
                                                                        };
                                                                    };
                                                                };



                                                                var NextColumnForResult = nextLetter(lastColName);
                                                                var NextColumnForShort = nextLetter(NextColumnForResult);
                                                                var NextColumnForDate = nextLetter(NextColumnForShort);
                                                                var rangeForResHead = ResultSheet.getRange(NextColumnForResult + 1 + ":" + NextColumnForDate + 1);
                                                                rangeForResHead.values = [["Result", "Short Links", "Date"]];
                                                                var toRangeLink = UrlItem.length + 1;
                                                                var range_Link = NextColumnForResult + 2 + ":" + NextColumnForDate + toRangeLink;
                                                                var rangeForResLink = ResultSheet.getRange(range_Link);


                                                                let data = FinalSheetSet;
                                                                var FROM = 1;
                                                                var TO = FROM + data.length - 1;
                                                                var RANEG = "A" + FROM.toString() + ":" + Aplhabet + TO.toString();
                                                                let range = ResultSheet.getRange(RANEG);
                                                                range.formulas = data;
                                                                range.format.autofitColumns();

                                                                var range_LinksRes = NextColumnForResult + 2 + ":" + NextColumnForResult + toRangeLink;
                                                                var rangeValOfLinks = ResultSheet.getRange(range_LinksRes);

                                                                rangeValOfLinks.format.wrapText = true;
                                                                rangeValOfLinks.format.columnWidth = 250;

                                                                ResultSheet.activate();

                                                                return context.sync()
                                                                    .then(function () {
                                                                        rangeForResLink.values = UrlItem;
                                                                        rangeForResLink.format.autofitColumns();

                                                                        //  AllSheetAutoFill();
                                                                        HideLoader();
                                                                    });


                                                            } else {




                                                                Excel.run(function (context) {

                                                                    let sheets = context.workbook.worksheets;

                                                                    let sheet = sheets.add("Result_" + ActiveSheet);
                                                                    sheet.load("name, position");

                                                                    return context.sync().then(function () {

                                                                        let ResultSheet = context.workbook.worksheets.getItem("Result_" + ActiveSheet);



                                                                        var HeadNames = $scope.UsedSheetValues[0];
                                                                        var markers = [];
                                                                        var lastColName;
                                                                        for (var n = 0; n < HeadNames.length; n++) {
                                                                            var Aplhabet = (n + 10).toString(36).toUpperCase();
                                                                            markers[i] = sheet.getRange(Aplhabet + 1);
                                                                            markers[i].values = HeadNames[n];
                                                                            if (n < HeadNames.length) {
                                                                                if (HeadNames[n] != "Result" && HeadNames[n] != "Short Links" && HeadNames[n] != "Date") {
                                                                                    lastColName = Aplhabet;
                                                                                };
                                                                            };
                                                                        };



                                                                        var NextColumnForResult = nextLetter(lastColName);
                                                                        var NextColumnForShort = nextLetter(NextColumnForResult);
                                                                        var NextColumnForDate = nextLetter(NextColumnForShort);
                                                                        var rangeForResHead = ResultSheet.getRange(NextColumnForResult + 1 + ":" + NextColumnForDate + 1);
                                                                        rangeForResHead.values = [["Result", "Short Links", "Date"]];
                                                                        var toRangeLink = UrlItem.length + 1;
                                                                        var range_Link = NextColumnForResult + 2 + ":" + NextColumnForDate + toRangeLink;
                                                                        var rangeForResLink = ResultSheet.getRange(range_Link);


                                                                        let data = FinalSheetSet;
                                                                        var FROM = 1;
                                                                        var TO = FROM + data.length - 1;
                                                                        var RANEG = "A" + FROM.toString() + ":" + Aplhabet + TO.toString();
                                                                        let range = ResultSheet.getRange(RANEG);
                                                                        range.formulas = data;
                                                                        range.format.autofitColumns();

                                                                        var range_LinksRes = NextColumnForResult + 2 + ":" + NextColumnForResult + toRangeLink;
                                                                        var rangeValOfLinks = ResultSheet.getRange(range_LinksRes);

                                                                        rangeValOfLinks.format.wrapText = true;
                                                                        rangeValOfLinks.format.columnWidth = 250;

                                                                        ResultSheet.activate();

                                                                        return context.sync()
                                                                            .then(function () {
                                                                                rangeForResLink.values = UrlItem;
                                                                                rangeForResLink.format.autofitColumns();

                                                                                //  AllSheetAutoFill();
                                                                                HideLoader();
                                                                            });





                                                                    });
                                                                });





                                                            };

                                                        });

                                                    });




                                                } else {
                                                    HideLoader();
                                                    LoadToast("Connection Issue. Please contact support@campaigntrackly.com");
                                                };


                                            };



                                        }).fail(function (error) {
                                            HideLoader();
                                            console.log(error);
                                            if (error.status != 200) {
                                                if (error.responseJSON.statusCode === 403 && error.responseJSON.message === "Expired token") {
                                                    RefreshToken(getFromLocal.refresh_token);
                                                    ShowLoader();
                                                    $scope.ApplyTemplate();
                                                }
                                                else {
                                                    LoadToast("Connection Issue. Please contact support@campaigntrackly.com");
                                                };
                                            } else {
                                                LoadToast("Connection Issue. Please contact support@campaigntrackly.com");

                                            }

                                            HideLoader();
                                            console.log(error);
                                        })

                                    };

                                });

                            });

                        });





















                    }

                })
            });



          

        };

        ////////////////////////change Date formate of excel ////////////////////////
        function getJsDateFromExcel(excelDateValue) {

            var d = new Date((excelDateValue - (25567 + 2)) * 86400 * 1000);
            month = '' + (d.getMonth() + 1),
                day = '' + d.getDate(),
                year = d.getFullYear();

            if (month.length < 2)
                month = '0' + month;
            if (day.length < 2)
                day = '0' + day;

            return [month, day, year].join('/');

        };




        ///////////////////////////////// Clear All Sheet /////////////////////////////////
        function ClearSheet() {
            Excel.run(function (context) {
                var worksheet = context.workbook.worksheets.getActiveWorksheet();
                var UsedFormularange = worksheet.getUsedRange();
                UsedFormularange.clear();
                return context.sync()
                    .then(function () {
                        // console.log("Clear Sheet")
                    })
            });
        };

        //////////////////////// contact to support ////////////////////////

        $scope.ContactSupport = function () {
            window.location.href = "mailto:support@campaigntrackly.com";
        };

        //////////////////////// Cehck user is logined or not ////////////////////////

        if (APIToken != null) {
            $scope.LoginDiv = true;
            $scope.MainPageDiv = false;
            $scope.NavBarDiv = false;
            $scope.StartedScreen = true;


            var isTokenExp = isTokenExpired(APIToken);

            if (isTokenExp) {
                ShowLoader();
               // console.log("Sesion Expired");
                RefreshToken(getFromLocal.refresh_token);
                ShowLoader();
                $scope.getTagTemplates();

            } else {
                $scope.getTagTemplates();
            };



            if (!$scope.$$phase) {
                $scope.$apply();
            };

        } else {
            if (FirstTime) {
                $scope.LoginDiv = true;
            } else {
                $scope.LoginDiv = false;
            };
            $scope.MainPageDiv = true;
            $scope.NavBarDiv = true;
            HideLoader();


            if (!$scope.$$phase) {
                $scope.$apply();
            };
        };


        //////////////////////// Refresh App ////////////////////////
        $scope.RefreshApp = function () {
            window.location.reload();
        };

        //////////////////////// logout ////////////////////////

        $scope.logOut = function () {
            $scope.LoginDiv = false;
            $scope.MainPageDiv = true;
            $scope.NavBarDiv = true;

            window.localStorage.removeItem("APIToken");
        };


    });
});