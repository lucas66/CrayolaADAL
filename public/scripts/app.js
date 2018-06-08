(function () {
    angular.module('app', [
      'ngRoute',
      'AdalAngular',
          'angular-loading-bar'
    ])
      .config(config);
    function config($routeProvider, $httpProvider, $locationProvider, adalAuthenticationServiceProvider, cfpLoadingBarProvider) {
        $routeProvider
                    .when('/', {
                        templateUrl: mainViewPath,
                        controller: 'MainController',
                        controllerAs: 'main',
                    })
                    .otherwise({
                        templateUrl: mainViewPath,
                        controller: 'MainController',
                        controllerAs: 'main',
                        requireADLogin: false,
                    });
        adalAuthenticationServiceProvider.init(
			{
			    instance: 'https://login.microsoftonline.com/',
			    clientId: enterpriseClientId,
			    anonymousEndpoints: ["/"],
			    requireADLogin: false,
			    endpoints: {
			        'https://graph.microsoft.com': 'https://graph.microsoft.com'
			    },
			    cacheLocation: 'localStorage'
			},
			$httpProvider
			);
        $locationProvider.html5Mode(true).hashPrefix('!');
        cfpLoadingBarProvider.includeSpinner = false;
    };
})();