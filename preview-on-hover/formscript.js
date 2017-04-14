"use strict";
/*global $*/
/*global _*/
/*global WebuiPopovers*/
/*global X2JS*/
/*global doT*/
/*global Xrm*/

var PreviewOnHover = {
    debug: false,
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
                return PreviewOnHover.buildEntityQuery(entityMetadata);
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

    getEntityMetadataFromCache: function (entityType) {
        var deferred = $.Deferred();

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
                    deferred.resolve(PreviewOnHover.Cache.get(entityType));  // returning the enityMetadata
                });;


        } else {
            deferred.resolve(entityMetadata);
        }

        return deferred.promise();
    },

    buildEntityQuery: function (entityMetadata) {
        var deferred = $.Deferred();
        var requests = Array();

        // rebuild query from array of expand items
        var expandlist = [];

        _.each(entityMetadata.queryExpandItems, function (value, key, list) {
            if (value.fieldtype == "Lookup" || value.fieldtype == "Owner" || value.fieldtype == "Customer") {
                requests.push(PreviewOnHover.getPrimaryNameAttribute(value.target, value.fieldname));
            }
        });

        if (requests.length > 0) {
            var defer = $.when.apply($, requests);
            defer.done(function () {
                // This is executed only after every ajax request has been completed
                $.each(arguments, function (index, responseData) {
                    // "responseData" will contain an array of response information for each specific request
                    expandlist.push(responseData.fieldname + "($select=" + responseData.PrimaryNameAttribute + ")");
                });
                if (expandlist.length > 0) {
                    entityMetadata.queryexpand = "";
                    entityMetadata.queryexpand = "$expand=" + expandlist.join(",");
                }
                deferred.resolve(entityMetadata);
            });
        } else {
            deferred.resolve(entityMetadata);
        }

        return deferred.promise();
    },

    retrieveEntityAndAttributeMetadata: function (entityType) {
        var deferred = $.Deferred();

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

                deferred.resolve();
            },
            error: function (data) {
                PreviewOnHover.log("retrieveEntityAndAttributeMetadata(" + entityType + ") - error: ");
                PreviewOnHover.log(data);
                $("#errorMessage").text(data.responseJSON.error.message);
                deferred.resolve();
            }
        });

        return deferred.promise();
    },

    retrieveFormMetadata: function (entityType) {
        var deferred = $.Deferred();

        var query = "systemforms?$filter=objecttypecode eq '" + entityType + "' and type eq 6";
        var url = PreviewOnHover.organizationURI + "/api/data/v8.0/" + query;

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
                                deferred.resolve();
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
                        "<p><h3>" + (entityMetadata.DisplayName).toUpperCase() + ": {{=it." + entityMetadata.PrimaryNameAttribute + "}}</h3>"
                        + "<h5>" + "No Quick View Form - please create one." + "</h5>"
                        + "<a href='javascript:void(0)' class='onhover-refresh' title='refresh cache'>Refresh</a></br>"
                        + "</p>"
                    });
                    deferred.resolve();
                }

            },
            error: function (data) {
                PreviewOnHover.log("retrieveFormMetadata(" + entityType + ") - error: ");
                PreviewOnHover.log(data);
                $("#errorMessage").text(data.responseJSON.error.message);
            }
        });

        return deferred.promise();

    },

    setForm: function (formid, entityType) {
        var deferred = $.Deferred();

        var entityMetadata = PreviewOnHover.Cache.get(entityType);

        if (!entityMetadata) {
            PreviewOnHover.log("retrieveFormMetadata(" + entityType + ") - ERROR - MetaData not found in Cache");
            deferred.resolve();
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
                            "<p><h3>" + (entityMetadata.DisplayName).toUpperCase() + ": {{=it." + entityMetadata.PrimaryNameAttribute + "}}</h3>"
                            + formHTML
                            + formOptions
                            + "<a href='javascript:void(0)' class='onhover-refresh' title='refresh cache'>Refresh</a></br>"
                            + "</p>"

                        });
                        deferred.resolve();
                    });


            } else {
                PreviewOnHover.log("FORM NOT FOUND: " + formid + "/" + entityType);
                deferred.resolve();
            }

        }

        return deferred.promise();
    },

    formEngine: {
        buildTemplate: function (jsonObj, entityType) {
            var deferred = $.Deferred();

            PreviewOnHover.formEngine.parseFormJSON(jsonObj, entityType)
                .then(function (string) {
                    deferred.resolve(string);
                });

            return deferred.promise();
        },

        parseFormJSON: function (jsonObj, entityType) {
            var deferred = $.Deferred();
            var result = "<table class='form'>";//"<form>";

            if (jsonObj.form) {
                PreviewOnHover.formEngine.addForm(jsonObj.form, entityType)
                    .then(function (string) {
                        result += string;
                        result += "</table>";//"</form>";
                        deferred.resolve(result);
                    });
            } else {
                result += "</table>";//"</form>";
                deferred.resolve(result);
            }


            return deferred.promise();
        },

        addForm: function (jsonObj, entityType) {
            var deferred = $.Deferred();
            var result = "";//"<tabs>";

            if (jsonObj.tabs) {
                PreviewOnHover.formEngine.addTabs(jsonObj.tabs, entityType)
                    .then(function (string) {
                        result += string;
                        result += "";//"</tabs>";
                        deferred.resolve(result);
                    });
            } else {
                result += "";//"</tabs>";
                deferred.resolve(result);
            }

            return deferred.promise();
        },

        addTabs: function (jsonObj, entityType) {
            var deferred = $.Deferred();
            var result = "";//"<tab>";

            if (jsonObj.tab) {
                PreviewOnHover.formEngine.addTab(jsonObj.tab, entityType)
                    .then(function (string) {
                        result += string;
                        result += "";//"</tab>"
                        deferred.resolve(result);
                    });
            } else {
                result += "";//"</tab>"
                deferred.resolve(result);
            }

            return deferred.promise();
        },

        addTab: function (jsonObj, entityType) {
            var deferred = $.Deferred();
            var result = "";//"<columns>";

            if (jsonObj.columns) {
                PreviewOnHover.formEngine.addColumns(jsonObj.columns, entityType)
                    .then(function (string) {
                        result += string;
                        result += "";//"</columns>"
                        deferred.resolve(result);
                    });
            } else {
                result += "";//"</columns>"
                deferred.resolve(result);
            }

            return deferred.promise();
        },

        addColumns: function (jsonObj, entityType) {
            var deferred = $.Deferred();
            var result = "";//"<column>";

            if (jsonObj.column) {
                PreviewOnHover.formEngine.addColumn(jsonObj.column, entityType)
                    .then(function (string) {
                        result += string;
                        result += "";//"</column>"
                        deferred.resolve(result);
                    });
            } else {
                result += "";//"</column>"
                deferred.resolve(result);
            }

            return deferred.promise();
        },

        addColumn: function (jsonObj, entityType) {
            var deferred = $.Deferred();
            var result = "";//"<sections>";

            if (jsonObj.sections) {
                PreviewOnHover.formEngine.addSections(jsonObj.sections, entityType)
                    .then(function (string) {
                        result += string;
                        result += "";//"</sections>"
                        deferred.resolve(result);
                    });
            } else {
                result += "";//"</sections>"
                deferred.resolve(result);
            }

            return deferred.promise();
        },
        addSections: function (jsonObj, entityType) {
            var deferred = $.Deferred();
            var result = "";//"<section>";

            if (jsonObj.section && !(_.isArray(jsonObj.section))) {
                PreviewOnHover.formEngine.addSection(jsonObj.section, entityType)
                    .then(function (string) {
                        result += string;
                        result += "";//"</section>"
                        deferred.resolve(result);
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
                    deferred.resolve(result);

                });

            } else {
                result += "";//"</section>"
                deferred.resolve(result);
            }

            return deferred.promise();
        },

        addSection: function (jsonObj, entityType) {
            var deferred = $.Deferred();
            var result = "";//"<rows>";

            if (jsonObj.rows) {
                PreviewOnHover.formEngine.addRows(jsonObj.rows, entityType)
                    .then(function (string) {
                        result += string;
                        result += "";//"</rows>"
                        deferred.resolve(result);
                    });
            } else {
                result += "";//"</rows>"
                deferred.resolve(result);
            }

            return deferred.promise();
        },

        addRows: function (jsonObj, entityType) {
            var deferred = $.Deferred();
            var requests = Array();

            var result = "";//"<row>";

            if (jsonObj.row) {
                _.each(jsonObj.row, function (row, index, list) {
                    // only adding row that are objects who have controls with datafieldnames
                    if (row && _.isObject(row) && row.cell && row.cell.control && row.cell.control._datafieldname) {
                        requests.push(PreviewOnHover.formEngine.addRow(row, entityType));
                    }
                });

                if (requests.length > 0) {
                    var defer = $.when.apply($, requests);

                    defer.done(function () {
                        // This is executed only after every ajax request has been completed
                        $.each(arguments, function (index, responseData) {
                            // "responseData" will contain an array of response information for each specific request
                            result += responseData;
                        });

                        result += "";//"</row>"
                        deferred.resolve(result);

                    });
                } else {
                    result += "";//"</row>"
                    deferred.resolve(result);
                }

            } else {
                result += "";//"</row>"
                deferred.resolve(result);
            }

            return deferred.promise();
        },

        addRow: function (jsonObj, entityType) {
            var deferred = $.Deferred();
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
                        deferred.resolve(result);
                    });

            } else {
                result += "</tr>"
                deferred.resolve(result);
            }

            return deferred.promise();
        },

        addCellControl: function (jsonObj, entityType) {
            var deferred = $.Deferred();
            var result = "<td class='control'>";

            if (jsonObj.control && jsonObj.control._datafieldname) {
                PreviewOnHover.formEngine.buildDotTemplateFieldAndQuery(jsonObj.control._datafieldname, entityType)
                    .then(function (string) {
                        result += string;
                        result += "</td>"
                        deferred.resolve(result);
                    });
            } else {
                result += "</td>"
                deferred.resolve(result);
            }

            return deferred.promise();
        },

        addCellLabels: function (jsonObj, entityType) {
            var deferred = $.Deferred();
            var result = "<td class='labels'>";

            if (jsonObj.labels) {
                PreviewOnHover.formEngine.addLabels(jsonObj.labels, entityType)
                    .then(function (string) {
                        result += string;
                        result += "</td>"
                        deferred.resolve(result);
                    });
            } else {
                result += "</td>"
                deferred.resolve(result);
            }

            return deferred.promise();
        },

        addLabels: function (jsonObj, entityType) {
            var deferred = $.Deferred();
            var result = "<label>";

            if (jsonObj.label) {
                result += jsonObj.label._description + ": ";
            }

            result += "</label>";
            deferred.resolve(result);
            return deferred.promise();
        },

        buildDotTemplateFieldAndQuery: function (fieldname, entityType) {
            var deferred = $.Deferred();

            var entityMetadata = PreviewOnHover.Cache.get(entityType);

            var attribute = _.findWhere(entityMetadata.Attributes, { LogicalName: fieldname });
            var result = "";

            // type = "Owner" ==> owninguser or owningteam
            if (attribute.AttributeType == "Owner") {

                PreviewOnHover.addToQueryExpandItems("owninguser", attribute.AttributeType, "systemuser", entityType)
                    .then(function () {
                        return PreviewOnHover.addToQueryExpandItems("owningteam", attribute.AttributeType, "team", entityType);
                    }).then(function () {
                        result = "{{?it.owninguser}} {{=it.owninguser.fullname}} {{?}}"
                            + "{{?it.owningteam}} {{=it.owningteam.name}} {{?}}";
                        deferred.resolve(result);
                    })

            } else if (attribute.AttributeType == "Customer") {

                PreviewOnHover.addToQueryExpandItems("parentcustomerid_account", attribute.AttributeType, "account", entityType)
                    .then(function () {
                        return PreviewOnHover.addToQueryExpandItems("parentcustomerid_contact", attribute.AttributeType, "contact", entityType);
                    }).then(function () {
                        result = "{{?it.parentcustomerid_account}} {{=it.parentcustomerid_account.name}} {{?}}"
                            + "{{?it.parentcustomerid_contact}} {{=it.parentcustomerid_contact.fullname}} {{?}}";
                        deferred.resolve(result);
                    });

                // type = "lookup" ==> field + ###NEED TO FIND the LOOKUP TYPE=Target ... and LOOKUP PRIMARYNAME  
            } else if (attribute.AttributeType == "Lookup") {
                PreviewOnHover.addToQueryExpandItems(attribute.LogicalName, attribute.AttributeType, attribute.Targets[0], entityType)
                    .then(function () {
                        result = "{{?it.primarycontactid}} {{=it.primarycontactid.fullname}} {{?}}";
                        deferred.resolve(result);
                    });

            } else {
                // no need to add to query items
                PreviewOnHover.addToQueryExpandItems(attribute.LogicalName, attribute.AttributeType, "", entityType)
                    .then(function () {
                        result = "{{=it." + fieldname + "}}";
                        deferred.resolve(result);
                    });

            }

            return deferred.promise();
        }

    },

    addToQueryExpandItems: function (fieldname, fieldtype, target, entityType) {
        var deferred = $.Deferred();

        var entityMetadata = PreviewOnHover.Cache.get(entityType);
        entityMetadata.queryExpandItems.push({
            fieldname: fieldname,
            fieldtype: fieldtype,
            target: target
        });

        deferred.resolve();
        return deferred.promise();
    },

    getPrimaryNameAttribute: function (entityType, fieldname) {
        var deferred = $.Deferred();

        // check if entityMetadata already exists
        PreviewOnHover.getEntityMetadataFromCache(entityType)
            .then(function (entityMetadata) {
                deferred.resolve({
                    entityMetadata: entityMetadata,
                    entityType: entityType,
                    PrimaryNameAttribute: entityMetadata.PrimaryNameAttribute,
                    fieldname: fieldname
                });
            });

        return deferred.promise();
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