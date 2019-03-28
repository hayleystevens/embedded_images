'use strict';

// Wrap everything in an anonymous function to avoid polluting the global namespace
(function () {
  /**
   * This extension collects the IDs of each workbook the user is interested in
   * and stores this information in settings when the popup is closed.
   */
  const WorksheetSettingsKey = 'selectedWorksheets';
    let selectedWorksheets = [];

  $(document).ready(function () {
    // The only difference between an extension in a dashboard and an extension 
    // running in a popup is that the popup extension must use the method
    // initializeDialogAsync instead of initializeAsync for initialization.
    // This has no affect on the development of the extension but is used internally.
    tableau.extensions.initializeDialogAsync().then(function (openPayload) {
      // The openPayload sent from the parent extension in this sample is the 
      // default time interval for the refreshes.  This could alternatively be stored
      // in settings, but is used in this sample to demonstrate open and close payloads.
      $('#interval').val(openPayload);
      $('#closeButton').click(closeDialog);

      let worksheets = tableau.extensions.dashboardContent.dashboard.worksheets ;
      let visibleWorksheets = [];
      selectedWorksheets = parseSettingsForActiveWorksheets();

      // Loop through datasources in this sheet and create a checkbox UI 
      // element for each one.  The existing settings are used to 
      // determine whether a datasource is checked by default or not.
      worksheets.forEach(function (worksheet) {
        worksheet.forEach(function (worksheet) {
            let isActive = (selectedWorksheets.indexOf(worksheet.id) >= 0);

            if (visibleWorksheets.indexOf(worksheet.id) < 0) {
              addWorksheetItemToUI(worksheet, isActive);
              visibleWorksheets.push(worksheet.id);
            }
          });
        });
      });
      });

  /**
   * Helper that parses the settings from the settings namesapce and 
   * returns a list of IDs of the datasources that were previously
   * selected by the user.
   */
  function parseSettingsForActiveWorksheets() {
    let activeWorksheetsIdList = [];
    let settings = tableau.extensions.settings.getAll();
    if (settings.selectedWorksheets) {
      activeWorksheetsIdList = JSON.parse(settings.selectedWorksheets);
    }

    return activeWorksheetsIdList;
  }

  /**
   * Helper that updates the internal storage of datasource IDs
   * any time a datasource checkbox item is toggled.
   */
  function updateWorksheetList(id) {
    let idIndex = selectedWorksheets.indexOf(id);
    if (idIndex < 0) {
      selectedWorksheets.push(id);
    } else {
      selectedWorksheets.splice(idIndex, 1);
    }
  }

  /**
   * UI helper that adds a checkbox item to the UI for a datasource.
   */
  function addWorksheetItemToUI(worksheet, isActive) {
    let containerDiv = $('<div />');

    $('<input />', {
      type: 'checkbox', /**Want this to be a single select drop down */
      id: worksheet.id,
      value: worksheet.name,
      checked: isActive,
      click: function() {updateWorksheetList(worksheet.id) }
    }).appendTo(containerDiv);

    $('<label />', {
      'for': worksheet.id,
      text: worksheet.name,
    }).appendTo(containerDiv);

    $('#worksheets').append(containerDiv);
  }

  /** Stores the selected worksheets IDs in the extension settings, closes the dialog, and sends a payload back to the parent. 
   */
  function closeDialog() {
    let currentSettings = tableau.extensions.settings.getAll();
    tableau.extensions.settings.set(WorksheetSettingsKey, JSON.stringify(selectedWorksheets));

    tableau.extensions.settings.saveAsync().then((newSavedSettings) => {
      tableau.extensions.ui.closeDialog($('#interval').val());
    });
  }
})();