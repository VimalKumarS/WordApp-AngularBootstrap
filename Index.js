/// <reference path="./angular.js" />
var WordService = angular.module("WordService", ['ngRoute',
        'someModuleP', 'someModuleS', 'someModuleF']);

angular.module('someModuleP', []).provider('exp', function () {
    var magicNumber = null;
    return {
        // internal configuration data; configured through setter function


        // configuration method for setting the magic number
        setMagicNumber: function (magicNumber) {
            this.magicNumber = magicNumber;
        },

        $get: function () {
            // use the magic number explicitly provided through "setMagicNumber" or
            // otherwise default to the injected "magicNumber" constant
            var toBeReturnedMagicNumber = this.magicNumber ;

            // return the service instance
            return {
                getMagicNumber: function () {
                    return toBeReturnedMagicNumber;
                }
            };
        }
    }
});
angular.module('someModuleS', []).service('helloWorldFromService', function () {
    this.sayHello = function () {
        return "Hello, World!";
    };
});
angular.module('someModuleF', []).factory('helloWorldFromFactory', function () {
    return {
        sayHello: function () {
            return "Hello, World!";
        }
    };
});

WordService.config(["expProvider", function (expProvider) {
    expProvider.setMagicNumber(99);
}]);

WordService.run(["exp",function (exp) {
    console.log(exp.getMagicNumber());
}]);

WordService.controller("mainCtrl", ["$scope", "helloWorldFromFactory", "helloWorldFromService","exp", function ($scope, helloWorldFromFactory, helloWorldFromService,exp) {
    var factory = helloWorldFromFactory;
    var service = helloWorldFromService;

    $scope.getContentControl = function () {
        var x = helloWorldFromFactory.sayHello();
        var y = exp.getMagicNumber();
        Word.run(function (context) {
            var contentControls = context.document.contentControls;

            // Queue a command to load the content controls collection.
            contentControls.load('type,parentContentControl');

            return context.sync().then(function () {

                if (contentControls.items.length === 0) {
                    console.log("There isn't a content control in this document.");
                } else {
                    console.log(contentControls.items);
                }

            });
        });
    }
}]);





//Office.initialize = function (reason) {
    angular.element(document).ready(function () {
        angular.bootstrap(document, ["WordService", "ngRoute", "ngResource"]);
    });
//};


//http://www.chaosm.net/blog/2014/07/27/load-angularjs-after-office-initialized/
//https://www.itunity.com/article/building-excel-2016-addin-angular-enhanced-officejs-2637
//http://stackoverflow.com/questions/33100354/office-js-callback-breaks-angular-controller
//https://github.com/ITUnity/dev/blob/master/HelloExcel2016/HelloExcel2016Web/App/App.js