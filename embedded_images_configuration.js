'use strict';

// Wrap everything in an anonymous function to avoid polluting the global namespace
(function () {
  /**
   * This extension collects the IDs of each workbook the user is interested in
   * and stores this information in settings when the popup is closed.
   */
 
   $(document).ready(function () {
    // The only difference between an extension in a dashboard and an extension 
    // running in a popup is that the popup extension must use the method
    // initializeDialogAsync instead of initializeAsync for initialization.
    // This has no affect on the development of the extension but is used internally.
    tableau.extensions.initializeDialogAsync().then(function (openPayload) {
      $('#closeButton').click(closeDialog);
      showChooseSheetDialog();

      function showChooseSheetDialog () {
        // Clear out the existing list of sheets
        $('#choose_sheet_buttons').empty();
    
        // Set the dashboard's name in the title
        const dashboardName = tableau.extensions.dashboardContent.dashboard.name;
        $('#choose_sheet_title').text(dashboardName);
    
        // The first step in choosing a sheet will be asking Tableau what sheets are available
        const worksheets = tableau.extensions.dashboardContent.dashboard.worksheets;
    
        // Next, we loop through all of these worksheets and add buttons for each one
        worksheets.forEach(function (worksheet) {
          // Declare our new button which contains the sheet name
          const button = createButton(worksheet.name);
    
          // Create an event handler for when this button is clicked
          button.click(function () {
            // Get the worksheet name and save it to settings.
            filteredColumns = [];
            const worksheetName = worksheet.name;
            tableau.extensions.settings.set('sheet', worksheetName);
            tableau.extensions.settings.saveAsync().then(function () {
              // Once the save has completed, close the dialog and show the data table for this worksheet
              $('#choose_sheet_dialog').modal('toggle');
              loadSelectedMarks(worksheetName);
            });
          });
    
          // Add our button to the list of worksheets to choose from
          $('#choose_sheet_buttons').append(button);
        });
    
        // Show the dialog
        $('#choose_sheet_dialog').modal('toggle');
      }
    
      function createButton (buttonTitle) {
        const button =
        $(`<button type='button' class='btn btn-default btn-block'>
          ${buttonTitle}
        </button>`);
    
        return button;
      }
    

  /** Stores the selected worksheets IDs in the extension settings, closes the dialog, and sends a payload back to the parent. 
   */
  function closeDialog() {
    let currentSettings = tableau.extensions.settings.getAll();
    tableau.extensions.settings.set(WorksheetSettingsKey, JSON.stringify(selectedWorksheets));

    tableau.extensions.settings.saveAsync().then((newSavedSettings) => {
      tableau.extensions.ui.closeDialog($("").val());
    });
  }
})();
