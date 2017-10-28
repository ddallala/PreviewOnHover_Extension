(function () {
    console.log("===============IN CONTENT 2===============")
    var cssRef = document.createElement('link');
    cssRef.setAttribute('rel', 'stylesheet');
    cssRef.setAttribute('type', 'text/css');
    cssRef.setAttribute('href', chrome.extension.getURL("preview-on-hover/style.css"));

    var formscriptTag = document.createElement('script');
    formscriptTag.setAttribute('type', 'text/javascript');
    formscriptTag.setAttribute('src', chrome.extension.getURL("preview-on-hover/formscript.js"));

    var mainBody = document.querySelectorAll('body[scroll=no]');
    var formBody = document.querySelectorAll('body.refresh-form');

    // ADDING CONTENT TO ENTITY FORMS
    if (formBody && formBody.length > 0) {
        console.log("---------in FORM----------");

        // underscore
        var underscoreTag = document.createElement('script');
        underscoreTag.setAttribute('type', 'text/javascript');
        underscoreTag.setAttribute('src', "https://cdnjs.cloudflare.com/ajax/libs/underscore.js/1.8.3/underscore-min.js");

        // dot Templates
        var dotTag = document.createElement('script');
        dotTag.setAttribute('type', 'text/javascript');
        dotTag.setAttribute('src', "https://cdnjs.cloudflare.com/ajax/libs/dot/1.1.0/doT.min.js");

        // WEbUI Popover
        var popoverTag = document.createElement('script');
        popoverTag.setAttribute('type', 'text/javascript');
        popoverTag.setAttribute('src', chrome.extension.getURL("preview-on-hover/js/jquery.webui-popover.js"));

        var popoverCSSTag = document.createElement('link');
        popoverCSSTag.setAttribute('rel', 'stylesheet');
        popoverCSSTag.setAttribute('href', "https://cdn.jsdelivr.net/jquery.webui-popover/2.1.15/jquery.webui-popover.min.css");       

        // XML2JSON
        var xml2jsonTag = document.createElement('script');
        xml2jsonTag.setAttribute('type', 'text/javascript');
        xml2jsonTag.setAttribute('src', chrome.extension.getURL("preview-on-hover/js/xml2json.js"));

        // injecting tags
        formBody[0].appendChild(popoverTag);
        formBody[0].appendChild(popoverCSSTag);
        formBody[0].appendChild(dotTag);
        formBody[0].appendChild(underscoreTag);
        formBody[0].appendChild(xml2jsonTag);
        formBody[0].appendChild(cssRef);
        formBody[0].appendChild(formscriptTag);
    }

})();
