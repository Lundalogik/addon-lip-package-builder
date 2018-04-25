// Initialize namespace object if not already initialized in other script files.
var app = app || {};

app.metadataloader = {

    "openExistingMetadata" : function(vm) {
        try {
            // First, clear old values
            vm.existingMetadata = undefined;
            vm.metadata_jsonUploaded(false);
            
            // Let the user select a file and try to get information from it
            var uploadedMetadata = $.parseJSON(lbs.common.executeVba('LIPPackageBuilder.OpenExistingMetadata'));
            if (uploadedMetadata !== {}) {
                vm.existingMetadata = new Metadata(uploadedMetadata);
                
                vm.metadata_jsonUploaded(true);

                vm.uniqueName(vm.existingMetadata.uniqueName);
                vm.displayName(vm.existingMetadata.displayName);
                vm.description(vm.existingMetadata.description);
            }
        }
        catch (e) {
            alert('app.metadataloader.openExistingMetadata(): ' + e);
        }
    }
};
