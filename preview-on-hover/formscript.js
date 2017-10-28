"use strict";
/*global $*/
/*global _*/
/*global WebuiPopovers*/
/*global X2JS*/
/*global doT*/
/*global Xrm*/

var PreviewOnHover = {
    debug: true,
    log: function () {
        if (PreviewOnHover.debug) console.log.apply(this, arguments);
    },
    initialize: function (organizationURI) {
        PreviewOnHover.organizationURI = organizationURI;
        PreviewOnHover.Cache.load();

        // adding data to form
        $("body").append($('<div id="preview-on-hover-dialog-bg"></div>'));
        $("body").append($('<div id="preview-on-hover-dialog"><h1>Preview On Hover Settings</h1><p>Select the forms to show on each Entity</p><div id="preview-on-hover-dialog-content"></div><div style="clear:both"></div><div id="preview-on-hover-dialog-button-container"><button class="preview-on-hover-dialog-button" id="preview-on-hover-dialog-ok">OK</button></div></div>'));

        // HANDLER for SIMPLE CONTROL LOOKUP
        $(document).on({
            mouseenter: function (e) {

                // get the target
                var target = $(e.currentTarget);
                var selector = "#" + $(e.currentTarget).attr("id");

                // extract the attribute id
                var attribute_id = target.attr("id").split("_")[0];

                // extract the attribute from the Form
                var attribute = Xrm.Page.getAttribute(attribute_id);
                if (attribute && attribute.getValue) {

                    // extract entity type, id from the attribute
                    var lookupItem = attribute.getValue()[0];
                    var entityType = lookupItem.entityType;
                    var RecordId = (lookupItem.id).slice(1, -1);

                    // show Popup
                    PreviewOnHover.UI.showPopup(selector, RecordId, entityType);

                }

            },
            mouseleave: function (e) {
                //stuff to do on mouse leave
                // PreviewOnHover.log("mouseleave")
                //   WebuiPopovers.hide($(e.target));
                //   WebuiPopovers.updateContent($(e.target),'some html or text');
            }
        }, '.ms-crm-Lookup-Item');

        // HANDLER FOR SUB-GRID LOOKUP
        /*
        $(document).on({
            mouseenter: function (e) {
                // get the target
                var target = $(e.currentTarget);
                var selector = "#" + $(e.currentTarget).attr("id").replace("{", "\\{").replace("}", "\\}");

                // field type - only supporting primaryfield for now and not lookup
                if (target.attr("id").split("_")[1] == "primaryField") {

                    var controlname = $($(selector).parents(".ms-crm-ListControl-Ex-Lite")[0]).attr("id");
                    var gridControl = Xrm.Page.ui.controls.get(controlname);
                    var entityType = gridControl.getEntityName();
                    var RecordId = (target.attr("id").split("_")[2]).slice(1, -1);

                    // show Popup
                    PreviewOnHover.UI.showPopup(selector, RecordId, entityType);
                }
            }
        }, '.ms-crm-List-Link')
        */


        // adding handler for select new form
        $(document).on({
            change: function (element) {
                var formid = $(element.target).children("option").filter(":selected").val();
                var entityType  = $(element.target).attr("id").split("_")[1];
                PreviewOnHover.setForm(formid, entityType);
            }
        }, '.onhover-form-selection');

        $(document).on({
            click: function (element) {
                PreviewOnHover.log("---Deleting local cache--");
                PreviewOnHover.log(PreviewOnHover.Cache);
                PreviewOnHover.Cache.purge();
                PreviewOnHover.log(PreviewOnHover.Cache);
            }
        }, '.preview-on-hover-dialog-refresh');
        

        // adding handler to launch settings
        $(document).on({
            click: function (element) {
                PreviewOnHover.UI.buildSettingsDialog();
            }
        }, '.show-settings-dialog');

        // adding handler for form switch from dialog
        $(document).on({
            change: function (element) {
                PreviewOnHover.log("switch form");
                var formid = $(element.target).children("option").filter(":selected").val();
                var entityType  = $(element.target).attr("id").split("_")[1];
                PreviewOnHover.setForm(formid, entityType);
            }
        }, '.preview-on-hover-dialog-form-selection');

        // adding handler for Popover enable
        $(document).on({
            change: function (element) {
                PreviewOnHover.log("switch checkbox");
                var enablePreviewOnHover = ($(element.target).attr('checked') == "checked");
                var entityType  = $(element.target).attr("id").split("_")[1];
                PreviewOnHover.setEnablePreviewOnHover(entityType,enablePreviewOnHover);
            }
        }, '.preview-on-hover-dialog-form-checkbox');

    },
    UI: {
        showPopup: function (selector, RecordId, entityType) {
            
                    // check if entityMetadata already exists
                    PreviewOnHover.getEntityMetadataFromCache(entityType)
                        .then(function (entityMetadata) {
                            return PreviewOnHover.buildEntityExpandQuery(entityMetadata);
                        })
                        .then(function (entityMetadata) {

                            // need to check if PrewviewOnHover is enabled
                            if(entityMetadata.enablePreviewOnHover){
                                var query = entityMetadata.LogicalCollectionName + "(" + RecordId + ")?" + entityMetadata.queryexpand;
                                var url = PreviewOnHover.organizationURI + "/api/data/v8.0/" + query;
                
                                PreviewOnHover.ajax({
                                    url: url,
                                    headers: {
                                        'Accept': "application/json",
                                        'Content-Type': 'application/json; charset=utf-8',
                                        'OData-MaxVersion': "4.0",
                                        'OData-Version': "4.0",
                                        'Prefer': 'odata.include-annotations="*"'
                                    },
                                    method: 'GET'}).then(function(data){
                                        PreviewOnHover.log('retrieveDataAndShowPopUp - succes: ');
                                        PreviewOnHover.log(data);
                
                                        // show popover
                                        WebuiPopovers.show(selector, {
                                            content: (doT.template(entityMetadata.template))(data),
                                            cache: false,
                                            closeable: true,
                                            onHide: function ($element) {
                                                var pop = WebuiPopovers.getPop(selector);
                                                pop.destroy();
                                            }, // callback after hide
                                            width: 400
                                        });
                
                                    });
                                    /*,
                                    error: function (data) {
                                        PreviewOnHover.log('retrieveDataAndShowPopUp - error: ');
                                        //alert(url);
                                        PreviewOnHover.log(data);
                                        $("#errorMessage").text(data.responseJSON.error.message);
                                    }*/
                            } else {
                                PreviewOnHover.log("Preview on Hover is not enabled on this entity: " + entityMetadata.DisplayName)
                            }
            
                        });
                },
                buildSettingsDialog: function () {
                    var dialog = new PreviewOnHover.UI.Dialog();
   
                    dialog.show();
                },
                Dialog: function () {

                    this.show = function () {
                        var $dialog = $("#preview-on-hover-dialog");

                        // building content
                        var $content = $dialog.find("#preview-on-hover-dialog-content");
                        $content.empty();
                        var $table = $("<table><thead><tr><th>Entity</th><th>Show Hover</th><th>Form</th></tr></thead><tbody></tbody></table>");
                        var $tablebody = $table.find("tbody");

                        PreviewOnHover.log(PreviewOnHover.Cache.database)

                        PreviewOnHover.Cache.database.forEach(function(entityMetadata){
                            var $tr = $("<tr/>");

                            $tr.append("<td>"+entityMetadata.DisplayName+"</td>");
                            var $checkbox = $("<input class='preview-on-hover-dialog-form-checkbox' id='checkbox_" + entityMetadata.entityType + "_" + (Math.floor(Math.random() * 1000000) + 1) + "' type='checkbox' />");
                            if(entityMetadata.enablePreviewOnHover) $checkbox.attr('checked', true);
                            $tr.append($("<td/>").append($checkbox));

                            // building options
                            if(entityMetadata.forms == undefined) $tr.append("Need to hover over entity to show up here")
                            else if(entityMetadata.forms.length == 0) $tr.append("<span style='color:red'>No Quick View form, please create one</span>")
                            else {
                                var selectHTML = "<select class='preview-on-hover-dialog-form-selection' id='select_" + entityMetadata.entityType + "_" + (Math.floor(Math.random() * 1000000) + 1) + "'>";
                                _.each(entityMetadata.forms, function (element, index, list) {
                                    selectHTML += "<option value='" + element.formid + "' " + (element.formid == entityMetadata.formId ? "selected" : "") + ">" + element.name + "</option>";
                                })
                                selectHTML += "</select>";
                                $tr.append($("<td/>").append(selectHTML));
                            }
                            $tablebody.append($tr);
                        })
                        $content.append($table);

                        $content.append("<a href='javacript:void()' class='preview-on-hover-dialog-refresh'>Force Refresh</a>")
    
                        $dialog.fadeIn(250);
                        $dialog.find("input").first().focus();

                        var $dialogBg = $("#preview-on-hover-dialog-bg");
                        $dialogBg.fadeIn(250);
    
                        $("#preview-on-hover-dialog-ok").unbind().click(
                            function (event) {
                                $dialog.fadeOut(250);
                                $dialogBg.fadeOut(250);
                            });
                    }
                } // end Dialog
    },

    Cache: {
        load: function () {
            // load from storage
            if (typeof (Storage) !== "undefined") {
                PreviewOnHover.Cache.database = JSON.parse(localStorage.getItem("PreviewOnHoverDBCache")) || [];
            } else {
                PreviewOnHover.Cache.database = [];
            }
        },
        purge: function () {
            PreviewOnHover.Cache.database = [];
            if (typeof (Storage) !== "undefined") {
                localStorage.removeItem("PreviewOnHoverDBCache");
            }
        },
        database: [],
        add: function (entityType, data) {
            var entityMetadata = this.get(entityType);
            data = _.extend(data, { entityType: entityType }); // adding entity type to data

            if (!entityMetadata) { // new data
                entityMetadata = data;
                this.database.push(data);
            } else { // does an update
                this.update(entityMetadata, data);
            }

            // update storage
            if (typeof (Storage) !== "undefined") {
                try {
                    localStorage.setItem("PreviewOnHoverDBCache", JSON.stringify(PreviewOnHover.Cache.database));
                } catch(err){
                    console.log(err)
                }
            }

            return entityMetadata;
        },
        update: function (entityType, data) {
            var entityMetadata = this.get(entityType);
            data = _.extend(data, { entityType: entityType }); // ensuring not changing type

            if (!entityMetadata) { // new add
                this.add(entityType, data);
            } else { // update
                _.extend(entityMetadata, data);

            }

            // update storage
            if (typeof (Storage) !== "undefined") {
                try {
                    localStorage.setItem("PreviewOnHoverDBCache", JSON.stringify(PreviewOnHover.Cache.database));
                } catch(err){
                    console.log(err)
                }
            }
        },
        get: function (entityType) {
            return _.findWhere(this.database, { entityType: entityType });
        }
    },

    getEntityMetadataFromCache1: async function (entityType) {

        // check if entityMetadata already exists
        var entityMetadata = PreviewOnHover.Cache.get(entityType);

        // if not found then generate new entityMetadata in Cache
        if (!(entityMetadata && entityMetadata.entityType)) {
            PreviewOnHover.log("getEntityMetadataFromCache - no entity metadata for:" + entityType);


            // get EntityCollectionName and all Attributes to use in querying for data
            await PreviewOnHover.retrieveEntityAndAttributeMetadata(entityType);
                    //  find the form metatdata and add it to the Cache
            await PreviewOnHover.retrieveFormMetadata(entityType);
            return PreviewOnHover.Cache.get(entityType);  // returning the entityMetadata
        } else {
            return entityMetadata;
        }

    },

    getEntityMetadataFromCache: async function (entityType) {

            // check if entityMetadata already exists
            var entityMetadata = PreviewOnHover.Cache.get(entityType);

            // if not found then generate new entityMetadata in Cache
            if (!(entityMetadata && entityMetadata.entityType)) {
                PreviewOnHover.log("getEntityMetadataFromCache - no entity metadata for:" + entityType);


                // get EntityCollectionName and all Attributes to use in querying for data
                await PreviewOnHover.retrieveEntityAndAttributeMetadata(entityType);
                //  find the form metatdata and add it to the Cache
                await PreviewOnHover.retrieveFormMetadata(entityType);

                return PreviewOnHover.Cache.get(entityType);  // returning the enityMetadata

            } else {
                return entityMetadata;
            }
    },

    buildEntityExpandQuery: async function (entityMetadata) {

        var asyncCalls = Array();

        // rebuild query from array of expand items
        var expandlist = [];

        _.each(entityMetadata.queryExpandItems, function (value, key, list) {
            if (value.fieldtype == "Lookup" || value.fieldtype == "Owner" || value.fieldtype == "Customer") {
                asyncCalls.push(PreviewOnHover.getPrimaryNameAttribute(value.target, value.fieldname));
            }
        });

        if (asyncCalls.length > 0) {

            var values = await Promise.all(asyncCalls);
            // This is executed only after every ajax request has been completed
            $.each(values, function (index, responseData) {
                // "responseData" will contain an array of response information for each specific request
                expandlist.push(responseData.fieldname + "($select=" + responseData.PrimaryNameAttribute + ")");
            });
            if (expandlist.length > 0) {
                entityMetadata.queryexpand = "";
                entityMetadata.queryexpand = "$expand=" + expandlist.join(",");
            }
            return entityMetadata;
        } else {
            return entityMetadata;
        }

    },

    retrieveEntityAndAttributeMetadata: async function (entityType) {
            // get quick view forms for entity

            // Getting EntityForm Data
            var data = await PreviewOnHover.ajax({
                url: PreviewOnHover.organizationURI + "/api/data/v8.0/" + "/EntityDefinitions(LogicalName='" + entityType + "')?$select=LogicalName,PrimaryNameAttribute,PrimaryIdAttribute,EntitySetName,SchemaName&$expand=Attributes",
                headers: {
                    'Accept': "application/json",
                    'Content-Type': 'application/json; charset=utf-8',
                    'OData-MaxVersion': "4.0",
                    'OData-Version': "4.0"
                },
                method: 'GET'});

            PreviewOnHover.log("--retrieveEntityAndAttributeMetadata(" + entityType + ")--success")
            PreviewOnHover.log(data);

            PreviewOnHover.Cache.add(entityType, {
                enablePreviewOnHover: true,
                DisplayName: data.SchemaName,
                LogicalCollectionName: data.EntitySetName,
                PrimaryIdAttribute: data.PrimaryIdAttribute,
                PrimaryNameAttribute: data.PrimaryNameAttribute,
                Attributes: data.Attributes,
                queryExpandItems: [],
                queryexpand: ""
            });

            return;
/*
                error: function (data) {
                    PreviewOnHover.log("retrieveEntityAndAttributeMetadata(" + entityType + ") - error: ");
                    PreviewOnHover.log(data);
                    $("#errorMessage").text(data.responseJSON.error.message);
                    resolve();
                }
*/

    },

    retrieveFormMetadata: async function (entityType) {

        var query = "systemforms?$filter=objecttypecode eq '" + entityType + "' and type eq 6";
        var url = PreviewOnHover.organizationURI + "/api/data/v8.0/" + query;

        // get quick view forms for entity

        // Getting EntityForm Data
        var data = await PreviewOnHover.ajax({
                                        url: url,
                                        headers: {
                                            'Accept': "application/json",
                                            'Content-Type': 'application/json; charset=utf-8',
                                            'OData-MaxVersion': "4.0",
                                            'OData-Version': "4.0"
                                        },
                                        method: 'GET'});

        PreviewOnHover.log("retrieveFormMetadata(" + entityType + ") - succes: ");
        PreviewOnHover.log(data);

        // if at least one quick view form
        if (data.value && data.value.length > 0) {

            var entityMetadata = PreviewOnHover.Cache.get(entityType);

            if (!entityMetadata) {
                PreviewOnHover.log("retrieveFormMetadata(" + entityType + ") - ERROR - MetaData not found in Cache")
            } else {

                // save forms to the entityMetadata
                PreviewOnHover.Cache.update(entityType, {
                    forms: data.value
                });

                // take the first view
                await PreviewOnHover.setForm(data.value[0].formid, entityType)
                return;
            }

        } else {
            PreviewOnHover.log("no quick view forms for: " + entityType);
            var entityMetadata = PreviewOnHover.Cache.get(entityType);
            // overwrites the entitymetadata template attribute
            PreviewOnHover.Cache.update(entityType, {
                forms: [],
                formName: "none",
                formId: "",
                formXML: "",
                formJSON: {},
                formHTML: "",
                template:
                "<p><span class='pop-over-title'>" + (entityMetadata.DisplayName).toUpperCase() + ": {{=it." + entityMetadata.PrimaryNameAttribute + "}}</span>"
                + "<h5>" + "No Quick View Form - please create one." + "</h5>"
                + "<a href='javascript:void(0)' class='show-settings-dialog' title='Show Preview on Hover Settings'>Settings</a>"
                + "</p>"
            });
            return;
        }
/*
                },
                error: function (data) {
                    PreviewOnHover.log("retrieveFormMetadata(" + entityType + ") - error: ");
                    PreviewOnHover.log(data);
                    $("#errorMessage").text(data.responseJSON.error.message);
                }
            });
            */

    },

    // helper function to turn jquery ajax into promise
    ajax: function(params) {
        return new Promise(function(resolve,reject){
            $.ajax({
                url: params.url,
                headers: params.headers,
                method: params.method,
                //   dataType: 'json',
                success: function (data) {
                    resolve(data)
                },
                error: function(data){
                    reject(data)
                }
            });
        });
    },

    setEnablePreviewOnHover: function(entityType, enabled){
        var entityMetadata = PreviewOnHover.Cache.get(entityType);
        if (!entityMetadata) {
            PreviewOnHover.log("retrieveFormMetadata(" + entityType + ") - ERROR - MetaData not found in Cache");
            return;
        } else {
            PreviewOnHover.Cache.update(entityType, {
                enablePreviewOnHover: enabled
            });
        }
    },

    setForm: async function (formid, entityType) {
        var entityMetadata = PreviewOnHover.Cache.get(entityType);
        
        if (!entityMetadata) {
            PreviewOnHover.log("retrieveFormMetadata(" + entityType + ") - ERROR - MetaData not found in Cache");
            return;
        } else {
            var form = _.findWhere(entityMetadata.forms, { formid: formid });
            if (form) {

                var quickviewXML = form.formxml;
                var x2js = new X2JS();
                var quickviewJSON = x2js.xml_str2json(quickviewXML);

                var formHTML = await PreviewOnHover.formEngine.buildTemplate(quickviewJSON, entityType);
                // building form options
                var formOptions = "<select class='onhover-form-selection' id='select_" + entityType + "_" + (Math.floor(Math.random() * 1000000) + 1) + "'>";
                _.each(entityMetadata.forms, function (element, index, list) {
                    formOptions += "<option value='" + element.formid + "' " + (element.formid == formid ? "selected" : "") + ">" + element.name + "</option>";
                })
                formOptions += "</select>";

                // overwrites the entitymetadata template attribute
                PreviewOnHover.Cache.update(entityType, {
                    formName: form.name,
                    formId: form.formid,
                    formXML: quickviewXML,
                    formJSON: quickviewJSON,
                    formHTML: formHTML,
                    template:
                    "<p><span class='pop-over-title'>" + (entityMetadata.DisplayName).toUpperCase() + ": {{=it." + entityMetadata.PrimaryNameAttribute + "}}</span>"
                    + formHTML
                    + "<span class='pop-over-form-options'>"
                    + formOptions
                    + "<a href='javascript:void(0)' class='show-settings-dialog' title='Show Preview on Hover Settings'>Settings</a>"
                    + "</span>"
                    + "</p>"

                });
                return;
            } else {
                PreviewOnHover.log("FORM NOT FOUND: " + formid + "/" + entityType);
                return;
            }

        }
    },

    formEngine: {
        buildTemplate: async function (jsonObj, entityType) {
            var string = await PreviewOnHover.formEngine.parseFormJSON(jsonObj, entityType);
            return string;
        },

        parseFormJSON: async function (jsonObj, entityType) {
                var result = "<table class='form'>";//"<form>";
                
                if (jsonObj.form) {
                    result += await PreviewOnHover.formEngine.addForm(jsonObj.form, entityType);
                    result += "</table>";//"</form>";
                    return result;
                } else {
                    result += "</table>";//"</form>";
                    return result;
                }
        },

        addForm: async function (jsonObj, entityType) {
                var result = "";//"<tabs>";
                
                if (jsonObj.tabs) {
                    result += await PreviewOnHover.formEngine.addTabs(jsonObj.tabs, entityType);
                    result += "";//"</tabs>";
                    return result;
                } else {
                    result += "";//"</tabs>";
                    return result;
                }
        },

        addTabs: async function (jsonObj, entityType) {
                var result = "";//"<tab>";

                if (jsonObj.tab) {
                    result += await PreviewOnHover.formEngine.addTab(jsonObj.tab, entityType);
                    result += "";//"</tab>"
                    return result;
                } else {
                    result += "";//"</tab>"
                    return result;
                }
        },

        addTab: async function (jsonObj, entityType) {
                var result = "";//"<columns>";

                if (jsonObj.columns) {
                    result += await PreviewOnHover.formEngine.addColumns(jsonObj.columns, entityType);
                    result += "";//"</columns>"
                    return result;
                } else {
                    result += "";//"</columns>"
                    return result;
                }
        },

        addColumns: function (jsonObj, entityType) {
            return new Promise(function(resolve,reject){
                var result = "";//"<column>";

                if (jsonObj.column) {
                    PreviewOnHover.formEngine.addColumn(jsonObj.column, entityType)
                        .then(function (string) {
                            result += string;
                            result += "";//"</column>"
                            resolve(result);
                        });
                } else {
                    result += "";//"</column>"
                    resolve(result);
                }
            });
        },

        addColumn: async function (jsonObj, entityType) {
                var result = "";//"<sections>";

                if (jsonObj.sections) {
                    result += await PreviewOnHover.formEngine.addSections(jsonObj.sections, entityType);
                    result += "";//"</sections>"
                    return result;
                } else {
                    result += "";//"</sections>"
                    return result;
                }
        },
        addSections: async function (jsonObj, entityType) {
                var result = "";//"<section>";

                if (jsonObj.section && !(_.isArray(jsonObj.section))) {
                    result += await PreviewOnHover.formEngine.addSection(jsonObj.section, entityType);
                    result += "";//"</section>"
                    return result;
                } else if (jsonObj.section && _.isArray(jsonObj.section)) {
                    var asyncCalls = Array();

                    _.each(jsonObj.section, function (element, index, list) {
                        asyncCalls.push(PreviewOnHover.formEngine.addSection(element, entityType))
                    });

                    var values = await Promise.all(asyncCalls);
                    $.each(values, function (index, responseData) {
                        // "responseData" will contain an array of response information for each specific request
                        result += responseData;
                    });
                    result += "";//"</row>"
                    return result;

                } else {
                    result += "";//"</section>"
                    return result;
                }
        },

        addSection: async function (jsonObj, entityType) {
                var result = "";//"<rows>";

                if (jsonObj.rows) {
                    var string = await PreviewOnHover.formEngine.addRows(jsonObj.rows, entityType)
                    result += string;
                    result += "";//"</rows>"
                    return result;
                } else {
                    result += "";//"</rows>"
                    return result;
                }
        },

        addRows: async function (jsonObj, entityType) {
                var asyncCalls = [];

                var result = "";//"<row>";

                if (jsonObj.row) {
                    _.each(jsonObj.row, function (row, index, list) {
                        // only adding row that are objects who have controls with datafieldnames
                        if (row && _.isObject(row) && row.cell && row.cell.control && row.cell.control._datafieldname) {
                            asyncCalls.push(PreviewOnHover.formEngine.addRow(row, entityType))
                        }
                    });

                    if (asyncCalls.length > 0) {
                        var values = await Promise.all(asyncCalls);
                        //if(err) console.log(err)
                        $.each(values, function (index, responseData) {
                            // "responseData" will contain an array of response information for each specific request
                            result += responseData;
                        });
                        result += "";//"</row>"
                        return result;

                    } else {
                        result += "";//"</row>"
                        return result;
                    }

                } else {
                    result += "";//"</row>"
                    return result;
                }
        },

        addRow: async function (jsonObj, entityType) {
                var result = "<tr class='cell'>";

                if (jsonObj.cell) {
                    result += await PreviewOnHover.formEngine.addCellLabels(jsonObj.cell, entityType);
                    result += await PreviewOnHover.formEngine.addCellControl(jsonObj.cell, entityType);
                    result += "</tr>"
                    return result;

                } else {
                    result += "</tr>"
                    return result;
                }
        },

        addCellControl: async function (jsonObj, entityType) {
                var result = "<td class='control'>";

                if (jsonObj.control && jsonObj.control._datafieldname) {
                    var string = await PreviewOnHover.formEngine.addFieldToTemplateAndExpand(jsonObj.control._datafieldname, entityType)
                    result += string;
                    result += "</td>"
                    return result;
                } else {
                    result += "</td>"
                    return result;
                }
        },

        addCellLabels: async function (jsonObj, entityType) {
                var result = "<td class='labels'>";

                if (jsonObj.labels) {
                    var string = await PreviewOnHover.formEngine.addLabels(jsonObj.labels, entityType)
                    result += string;
                    result += "</td>"
                    return result;
                } else {
                    result += "</td>"
                    return result;
                }
        },

        addLabels: async function (jsonObj, entityType) {
                var result = "<label>";

                if (jsonObj.label) {
                    result += jsonObj.label._description + ": ";
                }

                result += "</label>";
                return result;
        },

        addFieldToTemplateAndExpand: async function (fieldname, entityType) {
            try {
                var entityMetadata = PreviewOnHover.Cache.get(entityType);

                var attribute = _.findWhere(entityMetadata.Attributes, { LogicalName: fieldname });
                var result = "";

                // type = "Owner" ==> owninguser or owningteam
                if (attribute.AttributeType == "Owner") {

                    await PreviewOnHover.addToQueryExpandItems("owninguser", attribute.AttributeType, "systemuser", entityType);
                    await PreviewOnHover.addToQueryExpandItems("owningteam", attribute.AttributeType, "team", entityType);
                    result = "{{?it.owninguser}}"
                        + "<a href='" + PreviewOnHover.organizationURI + "/main.aspx?etn=systemuser&id={{=it.owninguser.ownerid}}&pagetype=entityrecord' target='_blank'>{{=it.owninguser.fullname}}</a>"
                        + "{{?}}"
                        + "{{?it.owningteam}}"
                        + "{{=it.owningteam.name}}"
                        + "<a href='" + PreviewOnHover.organizationURI + "/main.aspx?etn=team&id={{=it.owningteam.teamid}}&pagetype=entityrecord' target='_blank'>{{=it.owningteam.name}}</a>"
                        + "{{?}}";
                    return result;

                } else if (attribute.AttributeType == "Customer") {
                    await PreviewOnHover.addToQueryExpandItems(fieldname + "_account", attribute.AttributeType, "account", entityType);
                    await PreviewOnHover.addToQueryExpandItems(fieldname + "_contact", attribute.AttributeType, "contact", entityType);
                    result = "{{?it." + fieldname + "_account}}"
                        + "<a href='" + PreviewOnHover.organizationURI + "/main.aspx?etn=account&id={{=it." + fieldname + "_account.accountid}}&pagetype=entityrecord' target='_blank'>{{=it." + fieldname + "_account.name}}</a>"
                        + "{{?}}"
                        + "{{?it." + fieldname + "_contact}}"
                        + "<a href='" + PreviewOnHover.organizationURI + "/main.aspx?etn=contact&id={{=it." + fieldname + "_contact.contactid}}&pagetype=entityrecord' target='_blank'>{{=it." + fieldname + "_contact.fullname}}</a>"
                        + "{{?}}";
                        return result;

                    // type = "lookup" ==> field + ###NEED TO FIND the LOOKUP TYPE=Target ... and LOOKUP PRIMARYNAME  
                } else if (attribute.AttributeType == "Lookup") {
                    await PreviewOnHover.addToQueryExpandItems(fieldname, attribute.AttributeType, attribute.Targets[0], entityType);
                    var fielddata = await PreviewOnHover.getPrimaryNameAttribute(attribute.Targets[0], fieldname);

                    // adding to template
                    result = "{{?it." + fieldname + "}}"
                        + "<a href='" + PreviewOnHover.organizationURI + "/main.aspx?etn=" + fielddata.entityType + "&id={{=it." + fieldname + "." + fielddata.PrimaryIdAttribute + "}}&pagetype=entityrecord' target='_blank'>{{=it." + fieldname + "." + fielddata.PrimaryNameAttribute + "}}</a>"
                        + "{{?}}";
                    return result;
                } else {
                    // no need to add to expand query items
                    result = "{{=it." + fieldname + "}}";
                    return result;
                }
            } catch(err){
                console.log(err)
            }
        }

    }, // end of formengine

    addToQueryExpandItems: async function (fieldname, fieldtype, target, entityType) {

            var entityMetadata = PreviewOnHover.Cache.get(entityType);
            entityMetadata.queryExpandItems.push({
                fieldname: fieldname,
                fieldtype: fieldtype,
                target: target
            });

            return;
    },

    getPrimaryNameAttribute: async function (entityType, fieldname) {

        // check if entityMetadata already exists
        var entityMetadata = await PreviewOnHover.getEntityMetadataFromCache(entityType);
        return {
            entityMetadata: entityMetadata,
            entityType: entityType,
            PrimaryNameAttribute: entityMetadata.PrimaryNameAttribute,
            PrimaryIdAttribute: entityMetadata.PrimaryIdAttribute,
            fieldname: fieldname
        };
    },

};

$(document).ready(function () {
    console.log("=======FORM SCRIPT========");
    try {
        PreviewOnHover.initialize(Xrm.Page.context.getClientUrl());
    } catch (e) {
        console.log("can't show pop-over");
        console.log(e)
    }
});
