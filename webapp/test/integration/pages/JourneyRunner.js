sap.ui.define([
    "sap/fe/test/JourneyRunner",
	"re/airline/test/integration/pages/AeroRegistryList",
	"re/airline/test/integration/pages/AeroRegistryObjectPage"
], function (JourneyRunner, AeroRegistryList, AeroRegistryObjectPage) {
    'use strict';

    var runner = new JourneyRunner({
        launchUrl: sap.ui.require.toUrl('re/airline') + '/test/flp.html#app-preview',
        pages: {
			onTheAeroRegistryList: AeroRegistryList,
			onTheAeroRegistryObjectPage: AeroRegistryObjectPage
        },
        async: true
    });

    return runner;
});

