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
                    PreviewOnHover.showPopup(selector, RecordId, entityType);

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
                    PreviewOnHover.showPopup(selector, RecordId, entityType);
                }
            }
        }, '.ms-crm-List-Link')
        */


        // adding handler for select new form
        $(document).on({
            change: function (element) {
                PreviewOnHover.log("change");
                PreviewOnHover.setForm($(element.target).children("option").filter(":selected").val(), $(element.target).attr("id").split("_")[0]);
            }
        }, '.onhover-form-selection');

        // adding handler for refresh
        $(document).on({
            click: function (element) {
                PreviewOnHover.log("refresh");
                $(element.target).after("close dialog");
                $(element.target).remove();
                PreviewOnHover.Cache.purge();
            }
        }, '.onhover-refresh');

    },

    showPopup: function (selector, RecordId, entityType) {

        // check if entityMetadata already exists
        PreviewOnHover.getEntityMetadataFromCache(entityType)
            .then(function (entityMetadata) {
                return PreviewOnHover.buildEntityExpandQuery(entityMetadata);
            })
            .then(function (entityMetadata) {

                var query = entityMetadata.LogicalCollectionName + "(" + RecordId + ")?" + entityMetadata.queryexpand;
                var url = PreviewOnHover.organizationURI + "/api/data/v8.0/" + query;

                $.ajax({
                    url: url,
                    headers: {
                        'Accept': "application/json",
                        'Content-Type': 'application/json; charset=utf-8',
                        'OData-MaxVersion': "4.0",
                        'OData-Version': "4.0",
                        'Prefer': 'odata.include-annotations="*"'
                    },
                    method: 'GET',
                    //   dataType: 'json',
                    success: function (data) {
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

                    },
                    error: function (data) {
                        PreviewOnHover.log('retrieveDataAndShowPopUp - error: ');
                        //alert(url);
                        PreviewOnHover.log(data);
                        $("#errorMessage").text(data.responseJSON.error.message);
                    }
                });

            });
    },
    Cache: {
        load: function () {
            // load from storage
            if (typeof (Storage) !== "undefined") {
                PreviewOnHover.Cache.database = JSON.parse(localStorage.getItem("PreviewOnHoverDBCache")) || [];
            } else {
                PreviewOnHover.Cache.database = [];
            }
            PreviewOnHover.Cache.database = []; // --- HACK
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
                localStorage.setItem("PreviewOnHoverDBCache", JSON.stringify(PreviewOnHover.Cache.database));
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
                localStorage.setItem("PreviewOnHoverDBCache", JSON.stringify(PreviewOnHover.Cache.database));
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
            PreviewOnHover.retrieveEntityAndAttributeMetadata(entityType)
                .then(function () {
                    //  find the form metatdata and add it to the Cache
                    return PreviewOnHover.retrieveFormMetadata(entityType)
                }).then(function () {
                    return PreviewOnHover.Cache.get(entityType);  // returning the entityMetadata
                });;


        } else {
            return entityMetadata;
        }

    },

    getEntityMetadataFromCache: function (entityType) {
        return new Promise(function(resolve, reject){

            // check if entityMetadata already exists
            var entityMetadata = PreviewOnHover.Cache.get(entityType);

            // if not found then generate new entityMetadata in Cache
            if (!(entityMetadata && entityMetadata.entityType)) {
                PreviewOnHover.log("getEntityMetadataFromCache - no entity metadata for:" + entityType);


                // get EntityCollectionName and all Attributes to use in querying for data
                PreviewOnHover.retrieveEntityAndAttributeMetadata(entityType)
                    .then(function () {
                        //  find the form metatdata and add it to the Cache
                        return PreviewOnHover.retrieveFormMetadata(entityType)
                    }).then(function () {
                        resolve(PreviewOnHover.Cache.get(entityType));  // returning the enityMetadata
                    });;


            } else {
                resolve(entityMetadata);
            }
        });
    },

    buildEntityExpandQuery: function (entityMetadata) {
        return new Promise(function(resolve, reject){
            var asyncCalls = Array();

            // rebuild query from array of expand items
            var expandlist = [];

            _.each(entityMetadata.queryExpandItems, function (value, key, list) {
                if (value.fieldtype == "Lookup" || value.fieldtype == "Owner" || value.fieldtype == "Customer") {
                    asyncCalls.push(PreviewOnHover.getPrimaryNameAttribute(value.target, value.fieldname));
                }
            });

            if (asyncCalls.length > 0) {
                var defer = $.when.apply($, asyncCalls);
                Promise.all(asyncCalls).then(function(values){
                    // This is executed only after every ajax request has been completed
                    $.each(values, function (index, responseData) {
                        // "responseData" will contain an array of response information for each specific request
                        expandlist.push(responseData.fieldname + "($select=" + responseData.PrimaryNameAttribute + ")");
                    });
                    if (expandlist.length > 0) {
                        entityMetadata.queryexpand = "";
                        entityMetadata.queryexpand = "$expand=" + expandlist.join(",");
                    }
                    resolve(entityMetadata);
                })
            } else {
                resolve(entityMetadata);
            }
        });
    },

    retrieveEntityAndAttributeMetadata: function (entityType) {

        return new Promise(function(resolve, reject){

            // get quick view forms for entity

            // Getting EntityForm Data
            $.ajax({
                url: PreviewOnHover.organizationURI + "/api/data/v8.0/" + "/EntityDefinitions(LogicalName='" + entityType + "')?$select=LogicalName,PrimaryNameAttribute,PrimaryIdAttribute,EntitySetName,SchemaName&$expand=Attributes",
                headers: {
                    'Accept': "application/json",
                    'Content-Type': 'application/json; charset=utf-8',
                    'OData-MaxVersion': "4.0",
                    'OData-Version': "4.0"
                },
                method: 'GET',
                //   dataType: 'json',
                success: function (data) {
                    PreviewOnHover.log("--retrieveEntityAndAttributeMetadata(" + entityType + ")--success")
                    PreviewOnHover.log(data);

                    PreviewOnHover.Cache.add(entityType, {
                        DisplayName: data.SchemaName,
                        LogicalCollectionName: data.EntitySetName,
                        PrimaryIdAttribute: data.PrimaryIdAttribute,
                        PrimaryNameAttribute: data.PrimaryNameAttribute,
                        Attributes: data.Attributes,
                        queryExpandItems: [],
                        queryexpand: ""
                    });

                    resolve();
                },
                error: function (data) {
                    PreviewOnHover.log("retrieveEntityAndAttributeMetadata(" + entityType + ") - error: ");
                    PreviewOnHover.log(data);
                    $("#errorMessage").text(data.responseJSON.error.message);
                    resolve();
                }
            });

        });//end promise
    },

    retrieveFormMetadata: function (entityType) {

        var query = "systemforms?$filter=objecttypecode eq '" + entityType + "' and type eq 6";
        var url = PreviewOnHover.organizationURI + "/api/data/v8.0/" + query;

        return new Promise(function(resolve,reject){
        // get quick view forms for entity

            // Getting EntityForm Data
            $.ajax({
                url: url,
                headers: {
                    'Accept': "application/json",
                    'Content-Type': 'application/json; charset=utf-8',
                    'OData-MaxVersion': "4.0",
                    'OData-Version': "4.0"
                },
                method: 'GET',
                //   dataType: 'json',
                success: function (data) {
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
                            PreviewOnHover.setForm(data.value[0].formid, entityType)
                                .then(function () {
                                    resolve();
                                });

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
                            + "<a href='javascript:void(0)' class='onhover-refresh' title='refresh cache'>Refresh</a></br>"
                            + "</p>"
                        });
                        resolve();
                    }

                },
                error: function (data) {
                    PreviewOnHover.log("retrieveFormMetadata(" + entityType + ") - error: ");
                    PreviewOnHover.log(data);
                    $("#errorMessage").text(data.responseJSON.error.message);
                }
            });

        }); // end promise

    },

    setForm: function (formid, entityType) {
        return new Promise(function(resolve,reject){
            var entityMetadata = PreviewOnHover.Cache.get(entityType);
            
            if (!entityMetadata) {
                PreviewOnHover.log("retrieveFormMetadata(" + entityType + ") - ERROR - MetaData not found in Cache");
                resolve();
            } else {
                var form = _.findWhere(entityMetadata.forms, { formid: formid });
                if (form) {
    
                    var quickviewXML = form.formxml;
                    var x2js = new X2JS();
                    var quickviewJSON = x2js.xml_str2json(quickviewXML);
    
                    PreviewOnHover.formEngine.buildTemplate(quickviewJSON, entityType)
                        .then(function (formHTML) {
                            // building form options
                            var formOptions = "<select class='onhover-form-selection' id='" + entityType + "_" + (Math.floor(Math.random() * 1000000) + 1) + "'>";
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
                                + "<a href='javascript:void(0)' class='onhover-refresh' title='refresh cache'>Refresh</a></br>"
                                + "</span>"
                                + "</p>"
    
                            });
                            resolve();
                        });
    
    
                } else {
                    PreviewOnHover.log("FORM NOT FOUND: " + formid + "/" + entityType);
                    resolve();
                }
    
            }
        });
    },

    formEngine: {
        buildTemplate: function (jsonObj, entityType) {
            return new Promise(function(resolve,reject){
                PreviewOnHover.formEngine.parseFormJSON(jsonObj, entityType)
                .then(function (string) {
                    resolve(string);
                });
            });
        },

        parseFormJSON: function (jsonObj, entityType) {
            return new Promise(function(resolve,reject){
                var result = "<table class='form'>";//"<form>";
                
                if (jsonObj.form) {
                    PreviewOnHover.formEngine.addForm(jsonObj.form, entityType)
                        .then(function (string) {
                            result += string;
                            result += "</table>";//"</form>";
                            resolve(result);
                        });
                } else {
                    result += "</table>";//"</form>";
                    resolve(result);
                }
            });

        },

        addForm: function (jsonObj, entityType) {
            return new Promise(function(resolve,reject){
                var result = "";//"<tabs>";
                
                if (jsonObj.tabs) {
                    PreviewOnHover.formEngine.addTabs(jsonObj.tabs, entityType)
                        .then(function (string) {
                            result += string;
                            result += "";//"</tabs>";
                            resolve(result);
                        });
                } else {
                    result += "";//"</tabs>";
                    resolve(result);
                }
            })
        },

        addTabs: function (jsonObj, entityType) {
            return new Promise(function(resolve,reject){
                var result = "";//"<tab>";

                if (jsonObj.tab) {
                    PreviewOnHover.formEngine.addTab(jsonObj.tab, entityType)
                        .then(function (string) {
                            result += string;
                            result += "";//"</tab>"
                            resolve(result);
                        });
                } else {
                    result += "";//"</tab>"
                    resolve(result);
                }
            });
        },

        addTab: function (jsonObj, entityType) {
            return new Promise(function(resolve,reject){
                var result = "";//"<columns>";

                if (jsonObj.columns) {
                    PreviewOnHover.formEngine.addColumns(jsonObj.columns, entityType)
                        .then(function (string) {
                            result += string;
                            result += "";//"</columns>"
                            resolve(result);
                        });
                } else {
                    result += "";//"</columns>"
                    resolve(result);
                }
            });
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

        addColumn: function (jsonObj, entityType) {
            return new Promise(function(resolve,reject){
                var result = "";//"<sections>";

                if (jsonObj.sections) {
                    PreviewOnHover.formEngine.addSections(jsonObj.sections, entityType)
                        .then(function (string) {
                            result += string;
                            result += "";//"</sections>"
                            resolve(result);
                        });
                } else {
                    result += "";//"</sections>"
                    resolve(result);
                }
            });
        },
        addSections: function (jsonObj, entityType) {
            return new Promise(function(resolve,reject){
                var result = "";//"<section>";

                if (jsonObj.section && !(_.isArray(jsonObj.section))) {
                    PreviewOnHover.formEngine.addSection(jsonObj.section, entityType)
                        .then(function (string) {
                            result += string;
                            result += "";//"</section>"
                            resolve(result);
                        });
                } else if (jsonObj.section && _.isArray(jsonObj.section)) {
                    var requests = Array();

                    _.each(jsonObj.section, function (element, index, list) {
                        requests.push(PreviewOnHover.formEngine.addSection(element, entityType))
                    });

                    var defer = $.when.apply($, requests);
                    defer.done(function () {
                        // This is executed only after every ajax request has been completed
                        $.each(arguments, function (index, responseData) {
                            // "responseData" will contain an array of response information for each specific request
                            result += responseData;
                        });

                        result += "";//"</row>"
                        resolve(result);

                    });

                } else {
                    result += "";//"</section>"
                    resolve(result);
                }
            });
        },

        addSection: function (jsonObj, entityType) {
            return new Promise(function(resolve,reject){
                var result = "";//"<rows>";

                if (jsonObj.rows) {
                    PreviewOnHover.formEngine.addRows(jsonObj.rows, entityType)
                        .then(function (string) {
                            result += string;
                            result += "";//"</rows>"
                            resolve(result);
                        });
                } else {
                    result += "";//"</rows>"
                    resolve(result);
                }
            });
        },

        addRows: function (jsonObj, entityType) {
            return new Promise(function(resolve,reject){
                var asyncCalls = [];

                var result = "";//"<row>";

                if (jsonObj.row) {
                    _.each(jsonObj.row, function (row, index, list) {
                        // only adding row that are objects who have controls with datafieldnames
                        if (row && _.isObject(row) && row.cell && row.cell.control && row.cell.control._datafieldname) {
                            asyncCalls.push(PreviewOnHover.formEngine.addRow(row, entityType))
                        }
                    });
                    console.log("++++++++++++++++++++ASYNC CALLS+++++++++++++++++++++++++:" + asyncCalls.length)
                    if (asyncCalls.length > 0) {
                        console.log("++++++++++++++++++++ERROR+++++++++++++++++++++++++");
                        Promise.all(asyncCalls).then(function(values) {
                            //if(err) console.log(err)
                            $.each(values, function (index, responseData) {
                                // "responseData" will contain an array of response information for each specific request
                                result += responseData;
                            });
                            result += "";//"</row>"
                            resolve(result);
                        });

                    } else {
                        result += "";//"</row>"
                        resolve(result);
                    }

                } else {
                    result += "";//"</row>"
                    resolve(result);
                }
            });
        },

        addRow: function (jsonObj, entityType) {
            return new Promise(function(resolve,reject){
                var result = "<tr class='cell'>";

                if (jsonObj.cell) {
                    PreviewOnHover.formEngine.addCellLabels(jsonObj.cell, entityType)
                        .then(function (string) {
                            result += string;
                            return PreviewOnHover.formEngine.addCellControl(jsonObj.cell, entityType);
                        })
                        .then(function (string) {
                            result += string;
                            result += "</tr>"
                            resolve(result);
                        });

                } else {
                    result += "</tr>"
                    resolve(result);
                }
            });
        },

        addCellControl: function (jsonObj, entityType) {
            return new Promise(function(resolve,reject){
                var result = "<td class='control'>";

                if (jsonObj.control && jsonObj.control._datafieldname) {
                    PreviewOnHover.formEngine.addFieldToTemplateAndExpand(jsonObj.control._datafieldname, entityType)
                        .then(function (string) {
                            result += string;
                            result += "</td>"
                            resolve(result);
                        });
                } else {
                    result += "</td>"
                    resolve(result);
                }
            });
        },

        addCellLabels: function (jsonObj, entityType) {
            return new Promise(function(resolve,reject){
                var result = "<td class='labels'>";

                if (jsonObj.labels) {
                    PreviewOnHover.formEngine.addLabels(jsonObj.labels, entityType)
                        .then(function (string) {
                            result += string;
                            result += "</td>"
                            resolve(result);
                        });
                } else {
                    result += "</td>"
                    resolve(result);
                }
            });
        },

        addLabels: function (jsonObj, entityType) {
            return new Promise(function(resolve,reject){
                var result = "<label>";

                if (jsonObj.label) {
                    result += jsonObj.label._description + ": ";
                }

                result += "</label>";
                resolve(result);
            });
        },

        addFieldToTemplateAndExpand: function (fieldname, entityType) {
            return new Promise(function(resolve,reject){

                var entityMetadata = PreviewOnHover.Cache.get(entityType);

                var attribute = _.findWhere(entityMetadata.Attributes, { LogicalName: fieldname });
                var result = "";

                // type = "Owner" ==> owninguser or owningteam
                if (attribute.AttributeType == "Owner") {

                    PreviewOnHover.addToQueryExpandItems("owninguser", attribute.AttributeType, "systemuser", entityType)
                        .then(function () {
                            return PreviewOnHover.addToQueryExpandItems("owningteam", attribute.AttributeType, "team", entityType);
                        }).then(function () {
                            result = "{{?it.owninguser}}"
                                + "<a href='" + PreviewOnHover.organizationURI + "/main.aspx?etn=systemuser&id={{=it.owninguser.ownerid}}&pagetype=entityrecord' target='_blank'>{{=it.owninguser.fullname}}</a>"
                                + "{{?}}"
                                + "{{?it.owningteam}}"
                                + "{{=it.owningteam.name}}"
                                + "<a href='" + PreviewOnHover.organizationURI + "/main.aspx?etn=team&id={{=it.owningteam.teamid}}&pagetype=entityrecord' target='_blank'>{{=it.owningteam.name}}</a>"
                                + "{{?}}";
                            resolve(result);
                        })

                } else if (attribute.AttributeType == "Customer") {
                    PreviewOnHover.addToQueryExpandItems(fieldname + "_account", attribute.AttributeType, "account", entityType)
                        .then(function () {
                            return PreviewOnHover.addToQueryExpandItems(fieldname + "_contact", attribute.AttributeType, "contact", entityType);
                        }).then(function () {
                            result = "{{?it." + fieldname + "_account}}"
                                + "<a href='" + PreviewOnHover.organizationURI + "/main.aspx?etn=account&id={{=it." + fieldname + "_account.accountid}}&pagetype=entityrecord' target='_blank'>{{=it." + fieldname + "_account.name}}</a>"
                                + "{{?}}"
                                + "{{?it." + fieldname + "_contact}}"
                                + "<a href='" + PreviewOnHover.organizationURI + "/main.aspx?etn=contact&id={{=it." + fieldname + "_contact.contactid}}&pagetype=entityrecord' target='_blank'>{{=it." + fieldname + "_contact.fullname}}</a>"
                                + "{{?}}";
                            resolve(result);
                        });

                    // type = "lookup" ==> field + ###NEED TO FIND the LOOKUP TYPE=Target ... and LOOKUP PRIMARYNAME  
                } else if (attribute.AttributeType == "Lookup") {
                    PreviewOnHover.addToQueryExpandItems(fieldname, attribute.AttributeType, attribute.Targets[0], entityType)
                        .then(function () {
                            return PreviewOnHover.getPrimaryNameAttribute(attribute.Targets[0], fieldname)
                        })
                        .then(function (fielddata) {
                            // adding to template
                            result = "{{?it." + fieldname + "}}"
                                + "<a href='" + PreviewOnHover.organizationURI + "/main.aspx?etn=" + fielddata.entityType + "&id={{=it." + fieldname + "." + fielddata.PrimaryIdAttribute + "}}&pagetype=entityrecord' target='_blank'>{{=it." + fieldname + "." + fielddata.PrimaryNameAttribute + "}}</a>"
                                + "{{?}}";
                            resolve(result);
                        });

                } else {
                    // no need to add to expand query items
                    result = "{{=it." + fieldname + "}}";
                    resolve(result);
                }
            });
        }

    },

    addToQueryExpandItems: function (fieldname, fieldtype, target, entityType) {
        return new Promise(function(resolve,reject){

            var entityMetadata = PreviewOnHover.Cache.get(entityType);
            entityMetadata.queryExpandItems.push({
                fieldname: fieldname,
                fieldtype: fieldtype,
                target: target
            });

            resolve();
        });
    },

    getPrimaryNameAttribute1: async function (entityType, fieldname) {

        // check if entityMetadata already exists
        var entityMetadata = await PreviewOnHover.getEntityMetadataFromCache(entityType);
        console.log("------AWAIT--------")
        console.log({
            entityMetadata: entityMetadata,
            entityType: entityType,
            PrimaryNameAttribute: entityMetadata.PrimaryNameAttribute,
            PrimaryIdAttribute: entityMetadata.PrimaryIdAttribute,
            fieldname: fieldname
        })
        return {
            entityMetadata: entityMetadata,
            entityType: entityType,
            PrimaryNameAttribute: entityMetadata.PrimaryNameAttribute,
            PrimaryIdAttribute: entityMetadata.PrimaryIdAttribute,
            fieldname: fieldname
        };
    },


    getPrimaryNameAttribute: function (entityType, fieldname) {
        return new Promise(function(resolve,reject){
            // check if entityMetadata already exists
            PreviewOnHover.getEntityMetadataFromCache(entityType)
                .then(function (entityMetadata) {
                    console.log("------PROMISE--------")
                    console.log({
                        entityMetadata: entityMetadata,
                        entityType: entityType,
                        PrimaryNameAttribute: entityMetadata.PrimaryNameAttribute,
                        PrimaryIdAttribute: entityMetadata.PrimaryIdAttribute,
                        fieldname: fieldname
                    })
                    resolve({
                        entityMetadata: entityMetadata,
                        entityType: entityType,
                        PrimaryNameAttribute: entityMetadata.PrimaryNameAttribute,
                        PrimaryIdAttribute: entityMetadata.PrimaryIdAttribute,
                        fieldname: fieldname
                    });
                });
        });
    }

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
