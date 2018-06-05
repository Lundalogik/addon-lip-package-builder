var vm = {};

// Load viewmodel
initPackageLoader = function(viewModel){
    vm = viewModel;
}

 parseExistingPackage = function(){
            //Clear all selected and 'inExistingPackage' properties for all objects
            try{
                clearCollection(vm.vbaComponents(), "selected");
                clearCollection(vm.sql(), "selected");
                ko.utils.arrayForEach(vm.tables(), function(table){
                    table.indeterminate(indeterminateStatus.NotSelected);
                    clearCollection(table.guiFields(),"selected");
                });
                clearCollection(vm.localizations(),"checked");
                clearCollection(vm.lbsApps(), 'checked');
                clearCollection(vm.actionpads(), 'checked');
            }
            catch(e){alert(e);}
            // Set details in "details" tab
            try{
                vm.uniqueName(vm.existingPackage.uniqueName);
            }
            catch(e){alert("Error parsing package: " + e);}
            // Flag objects loaded from the package
            loadExistingTables();
            loadExistingVba();
            loadExistingSQL();
            loadExistingLocalizations();
            loadExistingLBSApps();
            loadExistingActionpads();
            
}
/**
 * Sets all items in a collection as unselected and not in an existing package
 * based on the second parameter
 * @param {Array} collection 
 * @param {String} selectedProperty
 */
clearCollection = function(collection, selectedProperty){
    
    ko.utils.arrayForEach(collection, function(selectableObject){
        selectableObject.inExistingPackage(false);
        if(selectableObject.hasOwnProperty(selectedProperty) && ko.isObservable(selectableObject[selectedProperty])){
            selectableObject[selectedProperty](false);
            selectableObject.inExistingPackage(false);
        }
    })
}
/**
 * Flags the viewmodels tables collection as inExistingPackage and selected
 * 
 */        
loadExistingTables = function(){
    try{
        var existingPackageTables = vm.existingPackage.install.tables;
        // There might not be any tables or fields in the package.
        if (!existingPackageTables){
            return;
        }
        //find all mutual tables
        ko.utils.arrayForEach(existingPackageTables, function(et){
            
           ko.utils.arrayForEach(vm.tables(), function(table){
               if(table.guiFields()){
                   ko.utils.arrayForEach(table.guiFields(),function(field){
                   //Flag fields and tables as existing in package
                   ko.utils.arrayForEach(et.fields,function(ef){
                       if(ef.name == field.name && et.name == table.name){
                           field.inExistingPackage(true);
                           field.selected(true);
                           table.inExistingPackage(true);
                       }
                   });
               });
               }
            //set selected or partially selected tables
            table.indeterminate(table.getIndeterminate());
           });
           
        });
    }
    catch(e){alert("Error loading tables from package: " + e);}
}
/**
 * Flags the viewmodels vbaComponents collection as inExistingPackage and selected
 * 
 */        
loadExistingVba = function(){
    var existingPackageVba = vm.existingPackage.install.vba;
    if(!existingPackageVba){
        return;
    }
    try{
        //Fetch vba from package
        ko.utils.arrayForEach(existingPackageVba,function(eVba){
            //Fetch vba named the same as in the existing package as in the current database, select it and mark as inExistingPackage
            ko.utils.arrayForEach(ko.utils.arrayFilter(vm.vbaComponents(),function (v){
                return v.name == eVba.name;
            }), function(vbaComponent){
                vbaComponent.inExistingPackage(true);
                vbaComponent.selected(true);
            });
        });
    }
    catch(e){alert(e);}
}
/**
 * Flags the viewmodels sql collection as inExistingPackage and selected
 * 
 */        
loadExistingSQL = function(){
    var existingPackageSQL = vm.existingPackage.install.sql;
    if(!existingPackageSQL){
        return;
    }
    try{
        //Fetch sql from package
        ko.utils.arrayForEach(existingPackageSQL, function(eSql){
            
            //Fetch sql named the same as in the existing package as in the current database, select it and mark as inExistingPackage
            ko.utils.arrayForEach(ko.utils.arrayFilter(vm.sql(),function (s){
                return s.name == eSql.name;
            }), function(sql){
                sql.inExistingPackage(true);
                sql.selected(true);
            });
        });
    }
    catch(e){alert(e);}
}
/**
 * Flags the viewmodels localizations collection as inExistingPackage and selected
 * 
 */        
loadExistingLocalizations = function(){
    var existingPackageLocalizations = vm.existingPackage.install.localize;
    
    if(!existingPackageLocalizations) {
        return;
    }

    try {
        //Fetch localizations from package
        ko.utils.arrayForEach(existingPackageLocalizations, function(eLocalizations){
            //Fetch localizations named the same as in the existing package as in the current database, select it and mark as inExistingPackage
            
            ko.utils.arrayForEach(vm.localizations(),function (l){
                
                if(l.owner  == eLocalizations.owner && l.code == eLocalizations.code){
                    l.checked(true);
                    l.inExistingPackage(true);
                }
            });
        });
    }
    catch(e) {
        alert(e);
    }
}

/**
 * Flags the viewmodel's LBS Apps collection as inExistingPackage and checked
 */
loadExistingLBSApps = function() {
    var existingPackageLBSApps = vm.existingPackage.install.apps;
    
    if (!existingPackageLBSApps) {
        return;
    }
    
    try {
        // Loop LBS apps from opened package.
        ko.utils.arrayForEach(existingPackageLBSApps, function(existingLBSApp) {
            // Loop LBS Apps from the current application. If they have the same name as the one in the opened package,
            // select it and mark as inExistingPackage.
            ko.utils.arrayForEach(vm.lbsApps(), function (currentApplicationLBSApps) {
                if(currentApplicationLBSApp.name  === existingLBSApp.name) {
                    currentApplicationLBSApp.checked(true);
                    currentApplicationLBSApp.inExistingPackage(true);
                }
            });
        });
    }
    catch(e) {
        alert(e);
    }
}

/**
 * Flags the viewmodels actionpads collection as inExistingPackage and checked
 * 
 */  
loadExistingActionpads = function() {
    var existingPackageActionpads = vm.existingPackage.install.actionpads;
    
    if(!existingPackageActionpads) {
        return;
    }
    
    try {
        // Loop actionpads from opened package.
        ko.utils.arrayForEach(existingPackageActionpads, function(existingActionpad) {
            // Loop actionpads from the current application. If they are for the same table as the one in the opened package,
            // select it and mark as inExistingPackage.
            ko.utils.arrayForEach(vm.actionpads(), function (currentApplicationActionpad){
                if(currentApplicationActionpad.tableName  === existingActionpad.tableName) {
                    currentApplicationActionpad.checked(true);
                    currentApplicationActionpad.inExistingPackage(true);
                }
            });
        });
    }
    catch(e) {
        alert(e);
    }
}
