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
      $('#closeButton').click(closeDialog);

      let dashboard = tableau.extensions.dashboardContent.dashboard;
      let visibleWorksheets = [];
      selectedWorksheets = parseSettingsForActiveWorksheets();

      // Loop through datasources in this sheet and create a checkbox UI 
      // element for each one.  The existing settings are used to 
      // determine whether a datasource is checked by default or not.
      dashboard.worksheets.forEach(function (worksheet) {
        worksheet.getDataSourcesAsync().then(function (datasources) {
          datasources.forEach(function (datasource) {
            let isActive = (selectedWorksheets.indexOf(datasource.id) >= 0);

            if (visibleWorksheets.indexOf(datasource.id) < 0) {
              addDataSourceItemToUI(datasource, isActive);
              visibleWorksheets.push(datasource.id);
            }
          });
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
  function addDataSourceItemToUI(datasource, isActive) {
    let containerDiv = $('<div />');

    $('<input />', {
      type: 'checkbox',
      id: datasource.id,
      value: datasource.name,
      checked: isActive,
      click: function() { updateWorksheetList(datasource.id) }
    }).appendTo(containerDiv);

    $('<label />', {
      'for': datasource.id,
      text: datasource.name,
    }).appendTo(containerDiv);

    $('#datasources').append(containerDiv);
  }

  /**
   * Stores the selected datasource IDs in the extension settings,
   * closes the dialog, and sends a payload back to the parent. 
   */
  function closeDialog() {
    let currentSettings = tableau.extensions.settings.getAll();
    tableau.extensions.settings.set(WorksheetSettingsKey, JSON.stringify(selectedWorksheets));

    tableau.extensions.settings.saveAsync().then((newSavedSettings) => {
      });
  }
})();
