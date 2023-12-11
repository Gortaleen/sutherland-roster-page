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

interface ListContactGroupsResponse
  extends GoogleAppsScript.People.Schema.ListContactGroupsResponse {}

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
  function getContactsArr(
    contactType: string,
    quotaUser: string,
    contactGroupsList: ListContactGroupsResponse
  ) {
    // Get contacts
    const contactGroup = contactGroupsList.contactGroups!.find(
      (contactGroup) => contactGroup.name === contactType
    );
    const resourceName = contactGroup!.resourceName;
    const maxMembers = People.ContactGroups!.get(resourceName!, {
      quotaUser,
    }).memberCount;
    const resourceNames = People.ContactGroups!.get(resourceName!, {
      maxMembers,
      quotaUser,
    }).memberResourceNames;

    return People.People!.getBatchGet({
      resourceNames,
      personFields: ["names"],
      quotaUser,
    }).responses;
  }

  function addContactsToDoc(
    contactType: string,
    contacts: Array<PersonResponse>,
    body: GoogleAppsScript.Document.Body,
    activeResourceNames: PersonResponse
  ) {
    const contactArr = contacts
      .filter(function (contact) {
        const name = contact.person!.resourceName;

        return Object.getOwnPropertyDescriptor(activeResourceNames, name!);
      })
      .map(function (contact) {
        return [
          contact.person!.names![0].displayNameLastFirst,
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
  }

  function main() {
    // https://developers.google.com/apps-script/reference/document/document
    const body = DocumentApp.getActiveDocument().getBody();
    const sessionUserID = Session.getActiveUser().getEmail();
    // https://developers.google.com/admin-sdk/reference-overview
    const customerId = AdminDirectory.Users!.get(sessionUserID).customerId;
    // Get officers
    const officers = <Officers>AdminDirectory.Users?.list({
      customer: customerId,
    })
      .users?.filter((user) => user.suspended === false)
      .reduce(function (accumulator, user) {
        const userID = user.primaryEmail!.split("@")[0];
        Object.defineProperty(accumulator, userID, {
          value: user.name!.fullName,
        });
        return accumulator;
      }, {});
    // Get contact groups (labels)
    const contactGroupsList = People.ContactGroups?.list();
    // Get active members
    const activeResourceNames = getContactsArr(
      "Active",
      sessionUserID,
      contactGroupsList!
    )!.reduce(function (accumulator, response) {
      Object.defineProperty(accumulator, response.person!.resourceName!, {
        value: true,
      });
      return accumulator;
    }, {});
    // Get pipers
    const pipers = getContactsArr("Piper", sessionUserID, contactGroupsList!);
    // Get drummers
    const drummers = getContactsArr(
      "Drummer",
      sessionUserID,
      contactGroupsList!
    );
    let rangeElement;
    let style = {};

    // clear doc
    while (body.getNumChildren() > 1) body.removeChild(body.getChild(0));
    body.clear();

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
    body.editAsText().appendText("\n\n");
    // Pipers
    addContactsToDoc("Piper", pipers!, body, activeResourceNames);
    // Drummers
    addContactsToDoc("Drummer", drummers!, body, activeResourceNames);

    // document formatting
    Object.defineProperty(style, DocumentApp.Attribute.BOLD, { value: true });
    Object.defineProperty(style, DocumentApp.Attribute.FONT_SIZE, {
      value: 12,
    });

    //
    rangeElement = body.findText("Officers");
    rangeElement.getElement().setAttributes(style);
    //
    rangeElement = body.findText("Pipers");
    rangeElement.getElement().setAttributes(style);
    //
    rangeElement = body.findText("Drummers");
    rangeElement.getElement().setAttributes(style);

    body.setMarginTop(0);
  }

  return { main };
})();
