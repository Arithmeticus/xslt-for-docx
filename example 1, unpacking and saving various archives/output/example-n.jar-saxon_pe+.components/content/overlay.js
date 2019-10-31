<_text>var helloworld = {&#xD;
  onLoad: function() {&#xD;
    // initialization code&#xD;
    this.initialized = true;&#xD;
    this.strings = document.getElementById("helloworld-strings");&#xD;
  },&#xD;
  onMenuItemCommand: function(e) {&#xD;
    var promptService = Components.classes["@mozilla.org/embedcomp/prompt-service;1"]&#xD;
                                  .getService(Components.interfaces.nsIPromptService);&#xD;
    promptService.alert(window, this.strings.getString("helloMessageTitle"),&#xD;
                                this.strings.getString("helloMessage"));&#xD;
  },&#xD;
&#xD;
};&#xD;
window.addEventListener("load", function(e) { helloworld.onLoad(e); }, false);&#xD;
</_text>