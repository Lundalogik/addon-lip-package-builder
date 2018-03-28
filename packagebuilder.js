packagebuilder = {
    vm: {},
    initialize: function(vm){
        this.vm = vm;
    },

    "serializePackage": function() {
        // Serialize selected tables and fields and combine with localization data
        var packageTables = [];
        var tables = [];
        var packageRelations = [];
        var relations = {};
        var sqlObjects = [];
        var localizations = [];
        var actionpads = [];
        if (vm.uniqueName() === '') {
            alert("You must specify a name. If you are creating a new add-on, it must be unique (check GitHub).");
            return;
        }
        
        try {
            // For each selected table

            $.each(vm.selectedTables(),function(i,table){
                var packageTable = {};
                
                //Clone the table object
                packageTable = jQuery.extend(true,{},table);

                // Fetch local names from table with same name
                var localNameTable  = vm.localNames.Tables.filter(function(t){
                    return t.name == table.name;
                })[0];

                // Set singular and plural local names for table
                packageTable.localname_singular = localNameTable.localname_singular;
                packageTable.localname_plural = localNameTable.localname_plural;

                var icon = vm.tableIcons().filter(function(ti){
                    return ti.table == table.name;
                })[0];

                if(icon != null){
                        packageTable.attributes.icon = icon.binarydata;
                }
                
                // For each selected field in current table
                var fields = [];
                var packageFields = [];

                var selectedFields = jQuery.extend(true,{},table.selectedFields());
                $.each(selectedFields,function(j,field){
                    // First, make sure that it is an array to make the code work properly. If there is only one field in the table, it is a single object instead of the array it should be.
                    if (!_.isArray(localNameTable.Fields)) {
                        localNameTable.Fields = [localNameTable.Fields];
                    }
                    // Fetch local names from field with same name from the other data source
                    var localNameField = localNameTable.Fields.filter(function(f){
                        return f.name == field.name;
                    })[0];
                    
                    //Clone the field
                    var packageField = jQuery.extend(true,{},field);

                    // Set local names for current field
                    packageField.localname = jQuery.extend(true,{},localNameField);

                    // Create relations
                    try{
                        if(field.attributes.fieldtype == "relation"){
                            //Lookup if relation already added
                            var existingRelation = relations[field.attributes.idrelation];

                            if(existingRelation == null || existingRelation == undefined){
                                var packageRelation = new Relation(field.attributes.idrelation,table.name, field.name);
                                relations[field.attributes.idrelation] = packageRelation;


                            }
                            else{
                                existingRelation.table2 = table.name;
                                existingRelation.field2 = field.name;
                            }
                        }
                    }
                    catch(e){
                        alert(e);
                    }

                    if(packageField.localname && packageField.localname.name){
                        delete packageField.localname.name;
                    }

                    if(packageField.localname && packageField.localname.order){
                        delete packageField.localname.order;
                    }

                    //The separator is added correctly as a property on a field, instead of localname.
                    if(packageField.localname && packageField.localname.separator){
                        packageField.separator = packageField.localname.separator;

                        delete packageField.localname.separator;

                    }

                    if(packageField.separator && packageField.separator.order){
                        delete packageField.separator.order;
                    }

                    if(packageField.localname && packageField.localname.option){
                        delete packageField.localname.option;
                    }

                    // Push field to fields
                    fields.push(packageField);


                });

                // Set fields to the current table
                packageTable.fields = fields;

                // Push table to tables
                packageTables.push(packageTable);
            });


            //Add relations as the package expects
            for(idrelation in relations){
                if(relations[idrelation].table2 != ""){
                    packageRelations.push({"table1": relations[idrelation].table1,
                                            "field1": relations[idrelation].field1,
                                            "table2": relations[idrelation].table2,
                                            "field2": relations[idrelation].field2
                                            })
                }

            }

            var packageRelationFields = [];
            //Fetch all relationfields in package
            var index;
            for(index = 0;index < packageTables.length; ++index){
                var j;
                for (j = 0;j <  packageTables[index].fields.length; j++){
                    var f = packageTables[index].fields[j];
                    if (f.attributes.fieldtype == "relation"){
                    packageRelationFields.push({ "name":packageTables[index].name + '.' + f.name, "remove": 1});
                    }
                }
            }

            //Check if field is existing in an relation
            for (index = 0;index < packageRelationFields.length; index++){
                var rf = packageRelationFields[index];
                var j;
                for (j = 0; j < packageRelations.length;j++){
                    var rel = packageRelations[j];
                    if (rel.table1 + '.' + rel.field1 == rf.name || rel.table2 + '.' + rel.field2 == rf.name){
                        rf.remove = 0;
                    }
                }
            }

            //remove unpaired relationfields
            $.each(packageRelationFields,function(i,relField){
                if(relField.remove == 1){
                    $.each(packageTables, function(j,packageTable){
                        if(packageTable.name == relField.name.substring(0, relField.name.indexOf("."))){
                            var indexOfObjectToRemove;
                            //find the field to remove
                            $.each(packageTable.fields, function(k, packageField){
                                if (packageField.name == relField.name.substring(relField.name.indexOf(".") + 1)){
                                    indexOfObjectToRemove = k;
                                }
                            });
                            //remove field from package
                            if(indexOfObjectToRemove >= 0){
                                packageTable.fields.splice(indexOfObjectToRemove,1);
                            }
                        }


                });
            }
            });

            $.each(vm.selectedSql(),function(i, sql){
                sqlObjects.push({"name": sql.name, "definition": vm.sqlDefinitions()[sql.name]})
            });



        }
        catch(e){
            alert(e);
        }

        try {
            var arrComponents = [];
            $.each(vm.selectedVbaComponents(), function(i, component){
                arrComponents.push({
                    "name": component.name,
                    "relPath": "vba\\" + component.name + component.extension()
                });
            });


            var arrLocalizations = [];
            $.each(vm.selectedLocalizations(), function(i, locale) {
                arrLocalizations.push({
                    'owner' : locale.owner,
                    'code' : locale.code,
                    'context' : locale.context,
                    'sv': locale.sv,
                    'en_us' : locale.en_us,
                    'fi' : locale.fi,
                    'no' : locale.no,
                    'da' : locale.da
                });
            });

            var arrActionpads = [];
            $.each(vm.selectedActionpads(), function(i, a) {
                arrActionpads.push({
                    'tableName': a.tableName,
                    'fileName': a.fileName,
                    'relPath': "actionpads\\" + a.fileName
                });
            });
            
            // Build package json from database structure
            var package_json = {
                "install" : {
                },
                "createdWithLIPPackageBuilderVersion" : vm.lipPackageBuilderVersion(),
                "uniqueName" : vm.uniqueName(),
                "addonOrPackageVersion" : vm.versionNumber()
            };

            var bSomethingToInstall = false;

            if(packageTables.length > 0){
                package_json.install.tables = packageTables;
                bSomethingToInstall = true;
            }
            if(packageRelations.length > 0){
                package_json.install.relations = packageRelations;
                bSomethingToInstall = true;
            }

            if(sqlObjects.length > 0){
                package_json.install.sql = sqlObjects
                bSomethingToInstall = true;
            }
            
            if(arrLocalizations.length > 0){
                package_json.install.localize = arrLocalizations;
                bSomethingToInstall = true;
            }

            if(arrActionpads.length > 0){
                package_json.install.actionpads = arrActionpads;
                bSomethingToInstall = true;
            }

            if(arrComponents.length > 0){
                package_json.install.vba = arrComponents;
                bSomethingToInstall = true;
            }
            //lbs.log.debug(JSON.stringify(package_json));
        }
        catch(e) {
            alert("Error serializing LIP Package:\n\n" + e);
        }

        if(bSomethingToInstall) {
            // Save using VBA Method
            try{
                // Build metadata json from information added by user
                var metadata_json = {
                    "uniqueName": vm.uniqueName(),
                    "displayName": vm.displayName(),
                    "description": vm.description(),
                    "cloudCompatible": false,       //##TODO: Add this to GUI
                    "tags": [                       //##TODO: Add this to GUI
                    ]
                };

                // Build readme json from information added by user
                var readme_md = {
                    "displayName": vm.displayName(),
                    "description": vm.description(),
                    "cloudCompatible": false           //##TODO: Add this to GUI
                };
                
                // Build changelog json from information added by user
                var changelog_md = {
                    "displayName": vm.displayName(),
                    "authors": vm.authors(),
                    "versionNumber": vm.versionNumber(),
                    "versionComment": vm.versionComment(),
                    "date": moment().format("YYYY-MM-DD")
                };
                
                //Base64 encode the entire strings, commas don't do well in VBA calls.
                lbs.common.executeVba('LIPPackageBuilder.CreatePackage,'
                                        + window.btoa(JSON.stringify(package_json))
                                        + ',' + window.btoa(JSON.stringify(metadata_json))
                                        + ',' + window.btoa(JSON.stringify(readme_md))
                                        + ',' + window.btoa(JSON.stringify(changelog_md))
                                    );

            }
            catch(e) {
                alert(e);
            }
        }
        else {
            alert("You haven't selected anything for your new package...");
        }
    }
}
