// Initialize namespace object if not already initialized in other script files.
var app = app || {};

app.changelogloader = {

    "openExistingChangelog" : function(vm) {
        try {
            // First, clear old values
            vm.existingChangelogVersion = new Version('');
            
            // Let the user select a file and try to get information from it
            var existingChangelogInfo = $.parseJSON(lbs.common.executeVba('LIPPackageBuilder.OpenExistingChangelog'));
            if (existingChangelogInfo !== {}) {
                vm.existingChangelogVersion = new Version(existingChangelogInfo.versionNumber);
                vm.existingChangelogVersion.authors(existingChangelogInfo.authors);
                vm.changelog_mdUploaded(true);
            }
            else {
                vm.changelog_mdUploaded(false);                
            }
        }
        catch (e) {
            alert('app.changelogloader.openExistingChangelog(): ' + e);
        }
    }
};
