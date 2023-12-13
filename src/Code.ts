/**
 * https://github.com/Gortaleen/sutherland-roster-page
 * https://github.com/google/clasp#readme
 * https://github.com/google/clasp/blob/master/docs/typescript.md
 * https://www.typescriptlang.org/docs/handbook/release-notes/typescript-2-0.html#non-null-assertion-operator
 * https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Operators/Nullish_coalescing
 * https://www.typescriptlang.org/tsconfig#strict
 * https://www.typescriptlang.org/tsconfig#alwaysStrict
 * https://typescript-eslint.io/getting-started
 */

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function updateRosterRun() {
  RosterUpdate.main();
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function updateRosterForceUpdate() {
  RosterUpdate.main(true);
}

interface ListConnectionsResponse
  extends GoogleAppsScript.People.Schema.ListConnectionsResponse {}

interface PersonResponse
  extends GoogleAppsScript.People.Schema.PersonResponse {}

interface Officers extends GoogleAppsScript.AdminDirectory.Schema.Users {
  "drum.sergeant": string;
  manager: string;
  pm: string;
  quartermaster: string;
  secretary: string;
  treasurer: string;
}

const RosterUpdate = (function () {
  function getOfficers(customerId: string): Officers {
    if (!AdminDirectory.Users) {
      throw "AdminDirectory.Users.list not available";
    }

    return <Officers>AdminDirectory.Users.list({
      customer: customerId,
    })
      .users?.filter(
        (user) => user.orgUnitPath === "/Officers" && user.suspended === false
      )
      .reduce(function (accumulator, user) {
        const userID = user.primaryEmail!.split("@")[0];
        Object.defineProperty(accumulator, userID, {
          value: user.name!.fullName,
        });

        return accumulator;
      }, {});
  }

  function getCustomerId(sessionUserID: string) {
    if (!AdminDirectory.Users) {
      throw "AdminDirectory.Users.get not available";
    }

    // https://developers.google.com/admin-sdk/reference-overview
    const customerId = AdminDirectory.Users.get(sessionUserID).customerId;

    if (!customerId) {
      throw "No customerId";
    }

    return customerId;
  }

  function checkContactsChanged(
    scriptProperties: GoogleAppsScript.Properties.Properties,
    quotaUser: string
  ) {
    // first check to see if any contacts have been added or deleted
    const connectionsSyncToken = scriptProperties.getProperty(
      "CONNECTIONS_SYNC_TOKEN"
    );
    let listConnectionsResponse: ListConnectionsResponse | undefined;

    try {
      // try...catch is used here as a workaround to refresh the syncToken.
      // syncTokens expire after seven days if there are no changes to the
      // contacts list.  An error is thrown by the list function when it is
      // called with an expired syncToken:
      // https://developers.google.com/people/api/rest/v1/people.connections/list#google.people.v1.PeopleService.ListConnections

      // https://developers.google.com/people/api/rest/v1/people.connections/list
      listConnectionsResponse = People.People?.Connections?.list("people/me", {
        personFields: ["names", "metadata"],
        requestSyncToken: true,
        syncToken: connectionsSyncToken,
        quotaUser,
      });
    } catch (err) {
      if (
        (err as Error).message ===
        "API call to people.people.connections.list failed with error: Sync token is expired. Clear local cache and retry call without the sync token."
      ) {
        // https://developers.google.com/people/api/rest/v1/people.connections/list
        listConnectionsResponse = People.People?.Connections?.list(
          "people/me",
          {
            personFields: ["names", "metadata"],
            requestSyncToken: true,
            syncToken: "",
            quotaUser,
          }
        );
      } else {
        throw err;
      }
    }

    if (!listConnectionsResponse) {
      throw "People.Connections.List not available";
    }
    if (listConnectionsResponse?.nextSyncToken) {
      scriptProperties.setProperty(
        "CONNECTIONS_SYNC_TOKEN",
        listConnectionsResponse.nextSyncToken || ""
      );
    }

    /*
    people.connections.list returns the number of contacts added, edited, or
    deleted while the current syncToken was valid.
    */
    return listConnectionsResponse.totalPeople;
  }

  function getBandMembers(
    scriptProperties: GoogleAppsScript.Properties.Properties,
    quotaUser: string
  ) {
    if (!People.ContactGroups) {
      throw "People.ContactGroups not available";
    }
    const listContactsGroupResponse = People.ContactGroups.list({ quotaUser });
    const maxMembers = listContactsGroupResponse.contactGroups?.find(
      (contactGroup) => contactGroup.name === "myContacts"
    )?.memberCount;
    const activeResourceName = scriptProperties.getProperty(
      "RESOURCE_NAME_ACTIVE"
    );
    const drummerResourceName = scriptProperties.getProperty(
      "RESOURCE_NAME_DRUMMER"
    );
    const piperResourceName = scriptProperties.getProperty(
      "RESOURCE_NAME_PIPER"
    );
    // https://developers.google.com/people/api/rest/v1/contactGroups/batchGet
    const batchGetContactGroupsResponse = People.ContactGroups.batchGet({
      maxMembers,
      resourceNames: [
        activeResourceName,
        drummerResourceName,
        piperResourceName,
      ],
      quotaUser,
    });
    const activeResourceNames = batchGetContactGroupsResponse.responses?.find(
      (response) => response.contactGroup?.resourceName === activeResourceName
    )?.contactGroup?.memberResourceNames;
    const activeDrummerResourceNames =
      batchGetContactGroupsResponse.responses
        ?.find(
          (response) =>
            response.contactGroup?.resourceName === drummerResourceName
        )
        ?.contactGroup?.memberResourceNames?.filter((resourceName) =>
          activeResourceNames?.includes(resourceName)
        ) || [];
    const activePiperResourceNames =
      batchGetContactGroupsResponse.responses
        ?.find(
          (response) =>
            response.contactGroup?.resourceName === piperResourceName
        )
        ?.contactGroup?.memberResourceNames?.filter((resourceName) =>
          activeResourceNames?.includes(resourceName)
        ) || [];
    // https://developers.google.com/people/api/rest/v1/people/getBatchGet
    const getPeopleResponse = People.People?.getBatchGet({
      personFields: "names,metadata",
      resourceNames: [
        ...activeDrummerResourceNames,
        ...activePiperResourceNames,
      ],
      quotaUser,
    });
    const drummers = getPeopleResponse?.responses?.filter((response) =>
      activeDrummerResourceNames.includes(response.person?.resourceName || "")
    );
    const pipers = getPeopleResponse?.responses?.filter((response) =>
      activePiperResourceNames.includes(response.person?.resourceName || "")
    );

    return [drummers, pipers];
  }

  function addOfficersToDoc(
    body: GoogleAppsScript.Document.Body,
    officers: Officers
  ) {
    // Add officer names to doc
    body.editAsText().appendText("Officers\n");
    body
      .appendListItem(
        "Pipe Major: " + (officers!.pm ? officers!.pm : "vacant") + "\n"
      )
      .setGlyphType(DocumentApp.GlyphType.BULLET);
    body
      .appendListItem(
        "Drum Sergeant: " +
          (officers["drum.sergeant"] ? officers["drum.sergeant"] : "vacant") +
          "\n"
      )
      .setGlyphType(DocumentApp.GlyphType.BULLET);
    body
      .appendListItem(
        "Manager: " + (officers.manager ? officers.manager : "vacant") + "\n"
      )
      .setGlyphType(DocumentApp.GlyphType.BULLET);
    body
      .appendListItem(
        "Secretary: " +
          (officers.secretary ? officers.secretary : "vacant") +
          "\n"
      )
      .setGlyphType(DocumentApp.GlyphType.BULLET);
    body
      .appendListItem(
        "Treasurer: " +
          (officers.treasurer ? officers.treasurer : "vacant") +
          "\n"
      )
      .setGlyphType(DocumentApp.GlyphType.BULLET);
    body
      .appendListItem(
        "Quartermaster: " +
          (officers.quartermaster ? officers.quartermaster : "vacant") +
          "\n"
      )
      .setGlyphType(DocumentApp.GlyphType.BULLET);

    return;
  }

  function addContactsToDoc(
    contactType: string,
    contacts: Array<PersonResponse> = [],
    body: GoogleAppsScript.Document.Body
  ) {
    const contactArr = contacts.map(function (contact) {
      return [
        contact.person!.names![0].displayNameLastFirst?.toUpperCase(),
        contact.person!.names![0].displayName,
      ];
    });

    body.editAsText().appendText(contactType + "s\n");

    contactArr.sort();
    contactArr.forEach((contact) =>
      body
        .appendListItem([contact[1]] + "\n")
        .setGlyphType(DocumentApp.GlyphType.BULLET)
    );
    body.editAsText().appendText("\n\n");

    return;
  }

  /**
   * Overwrite edits manually made to the Roster Google document
   * This code assumes the DriveActivity timestamp occurred shortly after the
   * LAST_UPDATED value filed when the script last updated the Roster doc.
   */
  function checkDocAltered(
    doc: GoogleAppsScript.Document.Document,
    scriptProperties: GoogleAppsScript.Properties.Properties
  ) {
    const queryDriveActivityResponse = DriveActivity.Activity?.query({
      pageSize: 1,
      itemName: "items/" + doc.getId(),
    });
    if (!queryDriveActivityResponse) {
      throw "DriveActivity.Activity not available";
    }
    const lastUpdatedByScriptStr =
      scriptProperties.getProperty("LAST_UPDATED") || "";
    if (!lastUpdatedByScriptStr) {
      return true;
    }
    const lastUpdatedByScriptDt = new Date(lastUpdatedByScriptStr);
    const lastAlteredStr =
      queryDriveActivityResponse.activities![0].timestamp || "";
    const lastAlteredDt = new Date(lastAlteredStr);
    // add one minute to offset the delay that occurs before DriveActivity timestamp files
    lastUpdatedByScriptDt.setMinutes(lastUpdatedByScriptDt.getMinutes() + 1);

    return lastAlteredDt > lastUpdatedByScriptDt;
  }

  function main(forceUpdate = false) {
    const quotaUser = Session.getActiveUser().getEmail();
    const customerId = getCustomerId(quotaUser);
    const scriptProperties = PropertiesService.getScriptProperties();
    const contactsChanged = checkContactsChanged(scriptProperties, quotaUser);
    const doc = DocumentApp.getActiveDocument();
    const docAltered = checkDocAltered(doc, scriptProperties);
    if (!contactsChanged && !forceUpdate && !docAltered) {
      // this assumes changes to Contacts will be made when changes are made to Officers.
      return;
    }
    const officers = getOfficers(customerId);
    const [drummers, pipers] = getBandMembers(scriptProperties, quotaUser);
    // https://developers.google.com/apps-script/reference/document/document
    const body = doc.getBody();
    let rangeElement;
    const style = { BOLD: true, FONT_SIZE: 12 };

    // clear doc
    while (body.getNumChildren() > 1) body.removeChild(body.getChild(0));
    body.clear();

    addOfficersToDoc(body, officers);
    body.editAsText().appendText("\n\n");
    addContactsToDoc("Piper", pipers, body);
    addContactsToDoc("Drummer", drummers, body);

    // document formatting
    rangeElement = body.findText("Officers");
    rangeElement.getElement().setAttributes(style);
    rangeElement = body.findText("Pipers");
    rangeElement.getElement().setAttributes(style);
    rangeElement = body.findText("Drummers");
    rangeElement.getElement().setAttributes(style);
    body.setMarginTop(0);
    scriptProperties.setProperty("LAST_UPDATED", new Date().toISOString());

    return;
  }

  return { main };
})();
