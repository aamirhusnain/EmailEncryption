var app = angular.module('MyAddin', ['ngMaterial'], function ($mdThemingProvider) {
    $mdThemingProvider.theme('default')
        .primaryPalette('teal', {
            'default': '500', // by default use shade 400 from the pink palette for primary intentions

        });
});
app.controller('AddinCtrl', function ($scope, $mdToast, $log) {

    LoaderShow();
    $scope.EncryptionMethod = "0";
    $scope.ExpirationPeriodDays = ["1", "2", "3", "4", "5", "6", "7", "10", "15", "20", "25", "30", "45", "60", "90", "120", "180"];

    var EncryptedMessage;
    var ExpirationDays = "0";


    function randomNumber() {
        var randomNumber = Math.floor(Math.random() * 100) + 1;
        //console.log(randomNumber);
    }

    randomNumber();

    // $scope.EncryptionMethod = "1";

    //$scope.newValue();


    //$scope.SecurePortalPage = false;
    //$scope.MainPage = true;
    //document.getElementById("1").textContent = "circle";
    //$scope.EasySecurePage = true;
    //$scope.RadioButtonPenal = true;
    //$scope.TLSVerifyPage = true;
    //$scope.UnprotectedPage = false;

    var PageNoo = 1;
    var ActivePageId = 1;
    var checkUser = window.localStorage.getItem("UserVisted");

    if (checkUser === null) {
        $scope.MainPage = true;
        $scope.SecurePortalPage = false;
        document.getElementById("1").textContent = "circle";
        $scope.EasySecurePage = true;
        $scope.RadioButtonPenal = false;
        $scope.TLSVerifyPage = true;
        $scope.UnprotectedPage = true;
        if (!$scope.$$phase) {
            $scope.$apply();
        }

        window.localStorage.setItem("UserVisted", true);
    }
    else {
        $scope.MainPage = false;
        $scope.SecurePortalPage = true;
        document.getElementById("1").textContent = "circle";
        $scope.EasySecurePage = true;
        $scope.TLSVerifyPage = true;
        $scope.UnprotectedPage = true;
        $scope.RadioButtonPenal = true;
        if (!$scope.$$phase) {
            $scope.$apply();
        }
    };


    function ClearRadio(PageNo) {
        document.getElementById("1").textContent = "radio_button_unchecked";
        document.getElementById("2").textContent = "radio_button_unchecked";
        document.getElementById("3").textContent = "radio_button_unchecked";
        document.getElementById("4").textContent = "radio_button_unchecked";
        document.getElementById(PageNo).textContent = "circle";
    };



    $scope.goToPage1 = function (pageNum) {
        ClearRadio(pageNum);
        ActivePageId = pageNum;
        $scope.SecurePortalPage = false;
        $scope.EasySecurePage = true;
        $scope.TLSVerifyPage = true;
        $scope.UnprotectedPage = true;
    };

    $scope.goToPage2 = function (pageNum) {
        ClearRadio(pageNum);
        ActivePageId = pageNum;
        $scope.SecurePortalPage = true;
        $scope.EasySecurePage = false;
        $scope.TLSVerifyPage = true;
        $scope.UnprotectedPage = true;
    };
    $scope.goToPage3 = function (pageNum) {
        ClearRadio(pageNum);
        ActivePageId = pageNum;
        $scope.SecurePortalPage = true;
        $scope.EasySecurePage = true;
        $scope.TLSVerifyPage = false;
        $scope.UnprotectedPage = true;
    };
    $scope.goToPage4 = function (pageNum) {
        ClearRadio(pageNum);
        ActivePageId = pageNum;
        $scope.SecurePortalPage = true;
        $scope.EasySecurePage = true;
        $scope.TLSVerifyPage = true;
        $scope.UnprotectedPage = false;
    };


    $scope.goToNextPage = function () {
        if (ActivePageId == 1) {
            $scope.goToPage2(2);
        }
        else if (ActivePageId == 2) {
            $scope.goToPage3(3);
        }
        else if (ActivePageId == 3) {
            $scope.goToPage4(4);
        } else {
            $scope.SkipAllPages();
        };

        // $scope.goToPage(PageNoo);
    };

    $scope.SkipAllPages = function () {
        $scope.SecurePortalPage = true;
        $scope.EasySecurePage = true;
        $scope.TLSVerifyPage = true;
        $scope.UnprotectedPage = true;
        $scope.RadioButtonPenal = true;
        $scope.MainPage = false;
    };



    Office.onReady(function () {

        //  console.log(Office.context.mailbox.item);

        $scope.newValue = function (value) {
            LoaderShow();
            var EncryptedMethod;

            RemoveAllHeaders();


            if (value == "1") {
                //Sending Email with secureportal
                EncryptedMethod = { "x-encryptmethod": "secureportal", "x-encryptplugin": "yes" }
                EncryptedMessage = "Encrypt Via Secure Portal";
               // $scope.SetExpirationDays("0");

            }
            else if (value == "2") {
                // Sending Email with Encrypt via Easy-Secure"
                EncryptedMethod = { "x-encryptmethod": "secureportal", "X-easy-secure": "Y", "x-encryptplugin": "yes" }
                EncryptedMessage = "Encrypt via Easy-Secure";
               // $scope.SetExpirationDays("0");
                //$scope.EncryptionConfirmSet("N");
                //$scope.pickupConfirmSet("N");

            } else if (value == "3") {
                //   Sending Email with TLS Verify
                EncryptedMethod = { "x-encryptmethod": "verifyopportunistic", "x-em-encrypt": "yes", "x-em-verification": "verify", "x-encryptplugin": "yes" }
                EncryptedMessage = "Encrypt with TLS Verify";
            }
            else {
                EncryptedMethod = undefined;
                LoaderHide();
            }

            if (EncryptedMethod) {
                Office.context.mailbox.item.internetHeaders.setAsync(
                    EncryptedMethod, setCallback);
            } else {
                RemoveAllHeaders();
            }

            function setCallback(asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    // console.log("Successfully set headers");
                    //   $scope.GetHeader();
                    LoaderHide();
                } else {
                    console.log("Error setting headers: " + JSON.stringify(asyncResult.error));
                };
            };

            // LoaderHide();

        };

        $scope.RequireMFAFun = function (selectedValMFA) {

            console.log(selectedValMFA);

            if (selectedValMFA == "authenticator,sms") {
                var RequireMFASelc = {
                    "X-Twofactor": selectedValMFA,
                };

                Office.context.mailbox.item.internetHeaders.setAsync(
                    RequireMFASelc, setCallbackOfMFA);
            };
        };

        function setCallbackOfMFA(asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                //console.log("Successfully set Days");
              //  LoaderHide();
            } else {
                LoaderHide();
                console.log("Error setting headers: " + JSON.stringify(asyncResult.error));
            };
        };


        $scope.SetExpirationDays = function (selectedDays) {

            var ExpirationDays = {
                "X-Portalexpire": selectedDays,
                "x-em-psk-expire-period": selectedDays
            };
            Office.context.mailbox.item.internetHeaders.setAsync(
                ExpirationDays, setCallbackOfDays);
        };

        //x - encryptmethod: secureportal
        //x - encryptplugin: yes
        //x - easy - secure: Y
        //x - portalexpire: 0
        //x - em - psk - expire - period: 0
        //x - readreceipt: Y
        //x - sendernotify: Y

        function setCallbackOfDays(asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                //console.log("Successfully set Days");
                LoaderHide();
            } else {
                LoaderHide();
                console.log("Error setting headers: " + JSON.stringify(asyncResult.error));
            };
        };

        $scope.pickupConfirmSet = function (PickValue) {
            var pickupConfirmtion = {
                "x-readreceipt": PickValue,
            };

            Office.context.mailbox.item.internetHeaders.setAsync(
                pickupConfirmtion, setCallbackOfDaysPicup);
        };


        $scope.EncryptionConfirmSet = function (EncryptConValue) {
            var encryptionConfirmtion = {
                "X-Sendernotify": EncryptConValue,
            };

            Office.context.mailbox.item.internetHeaders.setAsync(
                encryptionConfirmtion, setCallbackOfDaysPicup);
        };


        function setCallbackOfDaysPicup(asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Successfully set Days");
                LoaderHide();
            } else {
                LoaderHide();
                console.log("Error setting headers: " + JSON.stringify(asyncResult.error));
            };
        };

        $scope.HelpCenter = function () {
            window.open('https://helpdesk.encrypttitan.com/support/solutions/articles/47001161933-how-to-send-a-encrypted-email-using-the-outlook-plugin', '_blank');

        };


        $scope.CloseAddin = function () {
            Office.context.ui.closeContainer()
            // Office.addin.hide();
        };


        function RemoveAllHeaders() {
            Office.context.mailbox.item.internetHeaders.removeAsync(
                ["x-encryptmethod", "X-easy-secure", "x-encryptplugin", "x-em-encrypt", "x-em-verification", "X-Portalexpire", "x-em-psk-expire-period", "x-readreceipt", "X-Sendernotify"],
                removeCallback);

        };

        function removeCallback(asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                // console.log("Successfully removed selected headers");
            } else {
                // console.log("Error removing selected headers: " + JSON.stringify(asyncResult.error));
            }
        };

        LoaderHide();


    });

    if (!$scope.$$phase) {
        $scope.$apply();
    }

    function LoaderShow() {
        document.getElementById("Loader").style.display = "block";
        document.getElementById("LoaderDiv").style.display = "block";
    }

    function LoaderHide() {
        document.getElementById("Loader").style.display = "none";
        document.getElementById("LoaderDiv").style.display = "none";
    }


    function loadToast(alertMessage) {
        // var el = document.querySelectorAll('#zoom');
        $mdToast.show(
            $mdToast.simple()
                .textContent(alertMessage)
                .position('bottom')
                .hideDelay(1000))
            .then(function () {
                $log.log('Toast dismissed.');
            }).catch(function () {
                $log.log('Toast failed or was forced to close early by another toast.');
            });
        if (!$scope.$$phase) {
            $scope.$apply();
        }
    };

});