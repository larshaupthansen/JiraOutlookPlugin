
(() => {
  // The initialize function must be run each time a new page is loaded
  Office.initialize = (reason) => {
   jQuery(document).ready(function(){
      if (window.location.search) {
        // Check if warning should be displayed
        var warn = this.getParameterByName('warn');
        if (warn) {
          $('.not-configured-warning').show();
        } else {
          // See if the config values were passed
          // If so, pre-populate the values
          var url = this.getParameterByName('jiraUrl');

          $('#jira-url').val(url);         
        }
      }

      // When the GitHub username changes,
      // try to load Gists
      $('#jira-url').on('change', function(){
      
       
      });

      // When the Done button is clicked, send the
      // values back to the caller as a serialized
      // object.
      $('#settings-done').on('click', function() {
        var settings:any= {};

        settings.jiraUrl = $('#jira-url').val();

      });
    });
  };

  function sendMessage(message) {
    Office.context.ui.messageParent(message);
  }

  function getParameterByName(name, url) {
    if (!url) {
      url = window.location.href;
    }
    name = name.replace(/[\[\]]/g, "\\$&");
    var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
      results = regex.exec(url);
    if (!results) return null;
    if (!results[2]) return '';
    return decodeURIComponent(results[2].replace(/\+/g, " "));
  }

  async function run() {
    
       
    
  }
})();
