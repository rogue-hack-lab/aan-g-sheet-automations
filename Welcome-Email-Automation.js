// the specific Adopt a Neighbor community used in the subject line
const nameOfCommunity = 'Ashland';
// The names of the organizing team for this particular community"
const organizers = 'Cathy, Tonya, Chuck, Blaire, Mica, and Dylan';
// The URL of the volunteer spreadsheet being accessed and modified
const volunteerSpreadsheetUrl = 'https://docs.google.com/spreadsheets/d/1vukLl3Ccrqx_ckIaK5KQKn5MjdUzpnwqT7wEu5Pjm5M/edit';
const nameOfVolunteerWorksheet = 'Volunteers';
// The URL of the main spreadsheet being accessed and modified
const neighborSpreadsheetUrl = 'https://docs.google.com/spreadsheets/d/1vukLl3Ccrqx_ckIaK5KQKn5MjdUzpnwqT7wEu5Pjm5M/edit';
const nameOfNeighborWorksheet = 'Neighbors';
// The email address that volunteers and neighbors should reply to
const replyTo = 'neighborhood.response.team@gmail.com';
// The subject line for the three email templates;
const subject = 'Welcome to Adopt a Neighbor ' + nameOfCommunity;
// setting runFirstRowTest to true will cause the script to only process the first row
const runFirstRowTest = false;

// simple function that makes feeds substitution values to the email template
const populateTemplateWithNameSubstitution = (template, name) => {
  const templateWithSubstitutions = template;
  templateWithSubstitutions.substitutions = {neighborName: name, organizers, nameOfCommunity};
  return templateWithSubstitutions.evaluate().getContent();
}

function sendWelcomeEmailToVolunteer() {
  // Number of the column keeping track of whether or not the row was processed and email sent, so that the script can be restarted if it doesn't complete.
  const emailSentColumn = 19;
  const sheet = SpreadsheetApp.openByUrl(volunteerSpreadsheetUrl).getSheetByName(nameOfVolunteerWorksheet);
  const startRow = 2;
  const lastRow = sheet.getLastRow();
  const dataRange = sheet.getRange(startRow, 1, lastRow - 1, emailSentColumn);
  const data = dataRange.getValues();
  const certifiedTemplate = HtmlService.createTemplateFromFile('template-certified-volunteers');
  const uncertifiedTemplate = HtmlService.createTemplateFromFile('template-uncertified-volunteers');
  for (var i = 0; i < data.length; ++i) {
    const row = data[i];
    const certified = row[1] || undefined;
    const name = row[2];
    const emailAddress = row[5].trim() || undefined;
    const emailSent = row[18] || undefined;
    if (!emailSent && emailAddress) {
      let body;
      if(certified && certified.toLowerCase() === 'yes') {
        body = populateTemplateWithNameSubstitution(certifiedTemplate, name);
      } else {
        body = populateTemplateWithNameSubstitution(uncertifiedTemplate, name);
      }
      const options = {htmlBody: body, replyTo};
      MailApp.sendEmail(emailAddress, subject, body, options);
      sheet.getRange(startRow + i, emailSentColumn).setValue(new Date().toISOString());
      SpreadsheetApp.flush();
      Utilities.sleep(1 * 1000);
    } else {
      Logger.log('Error. Possible invalid data in:', row);
    }
    if(runFirstRowTest){
      break;
    }
  }
}

function sendWelcomeEmailToNeighbor() {
  // Number of the column keeping track of whether or not the row was processed and email sent, so that the script can be restarted if it doesn't complete.
  const emailSentColumn = 17;
  const sheet = SpreadsheetApp.openByUrl(neighborSpreadsheetUrl).getSheetByName(nameOfNeighborWorksheet);
  const startRow = 2;
  const lastRow = sheet.getLastRow();
  const dataRange = sheet.getRange(startRow, 1, lastRow - 1, emailSentColumn);
  const data = dataRange.getValues();
  const template = HtmlService.createTemplateFromFile('template-neighbors');
  for (var i = 0; i < data.length; ++i) {
    const row = data[i];
    const name = row[1];
    const emailAddress = row[4].trim() || undefined;
    const emailSent = row[16] || undefined;
    if(!emailSent && emailAddress) {
      const body = populateTemplateWithNameSubstitution(template, name);
      const options = {htmlBody: body, replyTo};
      MailApp.sendEmail(emailAddress, subject, body, options);
      sheet.getRange(startRow + i, emailSentColumn).setValue(new Date().toISOString());
      SpreadsheetApp.flush();
      Utilities.sleep(1 * 1000);
    } else if (!emailAddress) {
      Logger.log('Error. Possible invalid data in:', {emailAddress, name, emailSent});
    }
    if(runFirstRowTest){
      break;
    }
  }
}
