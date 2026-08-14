
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Job Application Form')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function submitJobApplication(formData, fileName, fileData) {
  try {
    const spreadsheetId = getSpreadsheetId();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Applications') || ss.insertSheet('Applications');
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const emailColumn = 5; 
      const existingEmails = sheet.getRange(2, emailColumn, lastRow - 1, 1).getValues();
      const emailExists = existingEmails.some(row => row[0] === formData.email);

      if (emailExists) {
        return {
          status: 'error',
          message: 'You have already submitted an application with this email address. Only one application per email is allowed.'
        };
      }
    }
    let cvUrl = '';
    if (fileData && fileName) {
      try {
        cvUrl = uploadCVToDrive(fileData, fileName, formData.email);
        Logger.log(`CV uploaded successfully: ${cvUrl}`);
      } catch (uploadError) {
        Logger.log(`Error uploading CV: ${uploadError.toString()}`);
        cvUrl = 'Upload Failed: ' + uploadError.message;
      }
    }
    const timestamp = new Date().toISOString();
    const row = [
      timestamp,
      formData.applyingFor,
      formData.firstName,
      formData.lastName,
      formData.email,
      formData.mobile,
      formData.dob,
      formData.religion,
      formData.maritalStatus,
      formData.gender,
      formData.nationality,
      formData.currentAddress,
      formData.permanentAddress,
      // SSC
      formData.sscInstitute,
      formData.sscBoard,
      formData.sscMajor,
      formData.sscYear,
      formData.sscGrade,
      formData.hscInstitute,
      formData.hscBoard,
      formData.hscMajor,
      formData.hscYear,
      formData.hscGrade,
      formData.bachelorDegree,
      formData.bachelorInstitute,
      formData.bachelorMajor,
      formData.bachelorGrade,
      formData.postgraduateDegree,
      formData.postgraduateInstitute,
      formData.postgraduateMajor,
      formData.postgraduateGrade,
      // Experience
      formData.yearsOfExperience,
      formData.lastOrganisation,
      formData.lastDesignation,
      formData.joiningDate,
      formData.endDate,
      formData.currentlyWorking ? 'Yes' : 'No',
      // CV URL
      cvUrl
    ];

    // Append to sheet
    sheet.appendRow(row);

    // Send confirmation email (optional)
    sendConfirmationEmail(formData.email, formData.firstName);

    return {
      status: 'success',
      message: 'Application submitted successfully!'
    };

  } catch (error) {
    console.error('Error submitting application:', error);
    return {
      status: 'error',
      message: 'Failed to submit application: ' + error.message
    };
  }
}

/**
 * Gets or initializes the active spreadsheet for storing applications
 * Uses the currently active spreadsheet instead of creating new ones
 * @returns {string} Spreadsheet ID
 */
function getSpreadsheetId() {
  try {
    // Use the active spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (!ss) {
      throw new Error('No active spreadsheet found. Please open a Google Sheet and run this script.');
    }

    const spreadsheetId = ss.getId();

    // Check if Applications sheet exists, if not create it
    let appSheet = ss.getSheetByName('Applications');

    if (!appSheet) {
      // Create Applications sheet in current spreadsheet
      appSheet = ss.insertSheet('Applications');

      // Set up headers
      const headers = [
        'Timestamp', 'Applying For',
        'First Name', 'Last Name', 'Email', 'Mobile', 'DOB', 'Religion',
        'Marital Status', 'Gender', 'Nationality', 'Current Address', 'Permanent Address',
        'SSC Institute', 'SSC Board', 'SSC Major', 'SSC Year', 'SSC Grade',
        'HSC Institute', 'HSC Board', 'HSC Major', 'HSC Year', 'HSC Grade',
        'Bachelor Degree', 'Bachelor Institute', 'Bachelor Major', 'Bachelor Grade',
        'Postgraduate Degree', 'Postgraduate Institute', 'Postgraduate Major', 'Postgraduate Grade',
        'Years of Experience', 'Last Organisation', 'Last Designation',
        'Joining Date', 'End Date', 'Currently Working', 'CV URL'
      ];
      appSheet.appendRow(headers);

      // Format header row
      appSheet.getRange(1, 1, 1, headers.length)
        .setBackground('#2e7d32')
        .setFontColor('#ffffff')
        .setFontWeight('bold');

      appSheet.setFrozenRows(1);
      Logger.log('Applications sheet created in active spreadsheet');
    }

    return spreadsheetId;
  } catch (error) {
    Logger.log('Error in getSpreadsheetId: ' + error.toString());
    throw error;
  }
}

/**
 * Fetches all job applications from the spreadsheet
 * @returns {Array} Array of application objects
 */
function getAllApplications() {
  try {
    // Use active spreadsheet and Applications sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Applications');

    if (!sheet) {
      Logger.log('Applications sheet not found');
      return [];
    }

    // Get all data (skip header row)
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('No applications found');
      return [];
    }

    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

    // Map data to objects
    const applications = data.map(row => ({
      timestamp: row[0],
      applyingFor: row[1],
      firstName: row[2],
      lastName: row[3],
      email: row[4],
      mobile: row[5],
      dob: row[6],
      religion: row[7],
      maritalStatus: row[8],
      gender: row[9],
      nationality: row[10],
      currentAddress: row[11],
      permanentAddress: row[12],
      // SSC Details
      sscInstitute: row[13],
      sscBoard: row[14],
      sscMajor: row[15],
      sscYear: row[16],
      sscGrade: row[17],
      // HSC Details
      hscInstitute: row[18],
      hscBoard: row[19],
      hscMajor: row[20],
      hscYear: row[21],
      hscGrade: row[22],
      // Bachelor Details
      bachelorDegree: row[23],
      bachelorInstitute: row[24],
      bachelorMajor: row[25],
      bachelorGrade: row[26],
      // Postgraduate Details
      postgraduateDegree: row[27],
      postgraduateInstitute: row[28],
      postgraduateMajor: row[29],
      postgraduateGrade: row[30],
      // Experience
      yearsOfExperience: row[31],
      lastOrganisation: row[32],
      lastDesignation: row[33],
      joiningDate: row[34],
      endDate: row[35],
      currentlyWorking: row[36],
      cvUrl: row[37]
    }));

    // Sort by timestamp (newest first)
    applications.sort((a, b) => {
      const dateA = new Date(a.timestamp);
      const dateB = new Date(b.timestamp);
      return dateB - dateA;
    });

    Logger.log(`Retrieved ${applications.length} applications`);
    return applications;

  } catch (error) {
    Logger.log('Error fetching applications: ' + error.toString());
    throw error;
  }
}

/**
 * Fetches all applications with their status from CandidateStatus sheet
 * @returns {Array} Array of application objects with status
 */
function getAllApplicationsWithStatus() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Get all applications
    const applications = getAllApplications();

    // Get or create CandidateStatus sheet
    let statusSheet = ss.getSheetByName('CandidateStatus');
    if (!statusSheet) {
      statusSheet = ss.insertSheet('CandidateStatus');
      statusSheet.getRange(1, 1, 1, 4).setValues([['Email', 'Status', 'Updated At', 'Updated By']]);
      statusSheet.getRange(1, 1, 1, 4)
        .setBackground('#2e7d32')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      statusSheet.setFrozenRows(1);
    }

    // Get all status data
    const lastRow = statusSheet.getLastRow();
    const statusMap = new Map();

    if (lastRow > 1) {
      const statusData = statusSheet.getRange(2, 1, lastRow - 1, 2).getValues();
      statusData.forEach(row => {
        if (row[0]) statusMap.set(row[0], row[1]);
      });
    }

    // Add status to each application
    const appsWithStatus = applications.map(app => ({
      ...app,
      status: statusMap.get(app.email) || 'pending'
    }));

    return appsWithStatus;

  } catch (error) {
    Logger.log('Error fetching applications with status: ' + error.toString());
    throw error;
  }
}

/**
 * Sends confirmation email to the applicant
 * @param {string} email - Applicant's email address
 * @param {string} firstName - Applicant's first name
 */
function sendConfirmationEmail(email, firstName) {
  try {
    const submissionTimeISO = new Date().toISOString();
    const submissionTimeReadable = formatISODate(submissionTimeISO);
    const subject = 'Job Application Received - Confirmation';
    const body = `
Dear ${firstName},

Thank you for submitting your job application. We have successfully received your application and our HR team will review it shortly.

Application Details:
- Submitted on: ${submissionTimeReadable}
- Reference: ${submissionTimeISO}

We will contact you if your profile matches our requirements. Please keep this email for your records.

Best regards,
HR Team
    `;

    MailApp.sendEmail(email, subject, body);
  } catch (error) {
    console.error('Error sending confirmation email:', error);
  }
}

/**
 * Uploads CV file to Google Drive ASSETS folder
 * @param {Object} fileData - The file data in base64 format
 * @param {string} fileName - The name of the file
 * @param {string} applicantEmail - Email of the applicant
 * @returns {string} File URL
 */
function uploadCVToDrive(fileData, fileName, applicantEmail) {
  try {
    // Get or create ASSETS folder
    const folder = getCVFolder();

    // Decode base64 and create file
    const blob = Utilities.newBlob(
      Utilities.base64Decode(fileData.split(',')[1]),
      fileData.split(';')[0].split(':')[1],
      fileName
    );

    // Create file with ISO timestamp and applicant email in name
    const sanitizedEmail = applicantEmail.replace(/[^a-zA-Z0-9]/g, '_');
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const newFileName = `${timestamp}_${sanitizedEmail}_${fileName}`;
    const file = folder.createFile(blob.setName(newFileName));

    // Set file description
    file.setDescription(`CV uploaded by ${applicantEmail} on ${new Date().toISOString()}`);

    Logger.log(`CV uploaded to ASSETS folder: ${newFileName}`);

    return file.getUrl();

  } catch (error) {
    Logger.log('Error uploading CV: ' + error.toString());
    throw error;
  }
}

/**
 * Gets or creates the ASSETS NAME folder in Google Drive
 * @returns {GoogleAppsScript.Drive.Folder} ASSETS NAME folder
 */
function getCVFolder() {
  const folderName = 'ASSETS NAME';
  const folders = DriveApp.getFoldersByName(folderName);

  if (folders.hasNext()) {
    Logger.log('ASSETS NAME folder found');
    return folders.next();
  } else {
    Logger.log('Creating ASSETS NAME folder in Google Drive');
    const folder = DriveApp.createFolder(folderName);

    // Set folder description
    folder.setDescription('Job Application Assets - CVs and Documents uploaded by applicants');

    return folder;
  }
}

/**
 * Updates the job description in Settings sheet
 * @param {string} newDescription - The new job description
 * @returns {Object} Response object
 */
function updateJobDescription(newDescription) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Get or create Settings sheet
    let settingsSheet = ss.getSheetByName('Settings');
    if (!settingsSheet) {
      settingsSheet = ss.insertSheet('Settings');
      settingsSheet.getRange(1, 1, 1, 2).setValues([['Setting Name', 'Value']]);
      settingsSheet.getRange(1, 1, 1, 2)
        .setBackground('#2e7d32')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      settingsSheet.setFrozenRows(1);
      settingsSheet.setColumnWidth(1, 200);
      settingsSheet.setColumnWidth(2, 800);
    }

    // Check if job description row exists
    const lastRow = settingsSheet.getLastRow();
    let jobDescRow = -1;

    if (lastRow > 1) {
      const settingNames = settingsSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < settingNames.length; i++) {
        if (settingNames[i][0] === 'JOB_DESCRIPTION') {
          jobDescRow = i + 2; // Add 2 because of header row and 0-based index
          break;
        }
      }
    }

    if (jobDescRow > 0) {
      // Update existing row
      settingsSheet.getRange(jobDescRow, 2).setValue(newDescription);
    } else {
      // Add new row
      settingsSheet.appendRow(['JOB_DESCRIPTION', newDescription]);
    }

    Logger.log('Job description updated in Settings sheet');
    return {
      status: 'success',
      message: 'Job description updated successfully'
    };
  } catch (error) {
    Logger.log('Error updating job description: ' + error.toString());
    return {
      status: 'error',
      message: 'Failed to update job description: ' + error.message
    };
  }
}

/**
 * Gets the current job description from Settings sheet
 * @returns {string} Current job description
 */
function getJobDescription() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName('Settings');

    if (!settingsSheet) {
      // Return default if Settings sheet doesn't exist
      return 'We are looking for a talented professional to join our team. The ideal candidate should have excellent skills and a passion for excellence.';
    }

    const lastRow = settingsSheet.getLastRow();
    if (lastRow <= 1) {
      return 'We are looking for a talented professional to join our team. The ideal candidate should have excellent skills and a passion for excellence.';
    }

    // Find job description row
    const data = settingsSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === 'JOB_DESCRIPTION') {
        return data[i][1] || 'We are looking for a talented professional to join our team. The ideal candidate should have excellent skills and a passion for excellence.';
      }
    }

    // Return default if not found
    return 'We are looking for a talented professional to join our team. The ideal candidate should have excellent skills and a passion for excellence.';
  } catch (error) {
    Logger.log('Error getting job description: ' + error.toString());
    return 'We are looking for a talented professional to join our team. The ideal candidate should have excellent skills and a passion for excellence.';
  }
}

/**
 * Formats ISO date string to human-readable format in Bangladesh Time (BST)
 * @param {string} isoString - ISO 8601 date string
 * @returns {string} Formatted date string
 */
function formatISODate(isoString) {
  try {
    const date = new Date(isoString);
    // Use Bangladesh Time (BST) - Asia/Dhaka timezone (UTC+6)
    return Utilities.formatDate(date, 'Asia/Dhaka', 'MMMM dd, yyyy HH:mm:ss z');
  } catch (error) {
    return isoString;
  }
}

/**
 * Authenticates admin user against Users sheet
 * @param {string} username - Admin username
 * @param {string} password - Admin password
 * @returns {Object} Authentication result
 */
function authenticateAdmin(username, password) {
  try {
    // Use active spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (!ss) {
      return {
        success: false,
        message: 'No active spreadsheet found'
      };
    }

    // Ensure Users sheet exists (creates with demo data if needed)
    getUsersSheetId();
    const sheet = ss.getSheetByName('Users');

    if (!sheet) {
      return {
        success: false,
        message: 'Users sheet not found'
      };
    }

    // Get all users data (skip header row)
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();

    // Find matching user
    for (let i = 0; i < data.length; i++) {
      const [storedUsername, storedPassword] = data[i];

      if (storedUsername === username && storedPassword === password) {
        // Log successful login with ISO timestamp
        const loginTime = new Date().toISOString();
        Logger.log(`Admin login successful - User: ${username}, Time: ${loginTime}`);

        return {
          success: true,
          user: {
            username: username,
            loginTime: loginTime
          }
        };
      }
    }

    // No match found
    Logger.log(`Failed login attempt for username: ${username}`);
    return {
      success: false,
      message: 'Invalid username or password'
    };

  } catch (error) {
    Logger.log('Authentication error: ' + error.toString());
    return {
      success: false,
      message: 'Authentication error: ' + error.message
    };
  }
}

/**
 * Gets or creates the Users sheet in the active spreadsheet
 * @returns {string} Spreadsheet ID
 */
function getUsersSheetId() {
  try {
    // Use the active spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (!ss) {
      throw new Error('No active spreadsheet found');
    }

    // Check if Users sheet exists
    let usersSheet = ss.getSheetByName('Users');

    if (!usersSheet) {
      // Create Users sheet in active spreadsheet
      usersSheet = ss.insertSheet('Users');

      // Set up headers
      usersSheet.getRange(1, 1, 1, 2).setValues([['Username', 'Password']]);

      // Add demo admin users
      const demoUsers = [
        ['admin', 'admin123'],
        ['hr_manager', 'hr2024'],
        ['recruiter', 'recruit@2024'],
        ['owner', 'owner@secure']
      ];
      usersSheet.getRange(2, 1, demoUsers.length, 2).setValues(demoUsers);

      // Format header row
      usersSheet.getRange(1, 1, 1, 2)
        .setBackground('#2e7d32')
        .setFontColor('#ffffff')
        .setFontWeight('bold');

      // Auto-resize columns
      usersSheet.autoResizeColumns(1, 2);
      usersSheet.setFrozenRows(1);

      Logger.log('Users sheet created with demo data in active spreadsheet');
    }

    return ss.getId();
  } catch (error) {
    Logger.log('Error in getUsersSheetId: ' + error.toString());
    throw error;
  }
}

/**
 * Updates candidate status in the CandidateStatus sheet
 * @param {string} email - Candidate's email
 * @param {string} newStatus - New status (pending/shortlisted/rejected)
 * @returns {Object} Response object
 */
function updateCandidateStatus(email, newStatus) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Get or create CandidateStatus sheet
    let statusSheet = ss.getSheetByName('CandidateStatus');
    if (!statusSheet) {
      statusSheet = ss.insertSheet('CandidateStatus');
      statusSheet.getRange(1, 1, 1, 4).setValues([['Email', 'Status', 'Updated At', 'Updated By']]);
      statusSheet.getRange(1, 1, 1, 4)
        .setBackground('#2e7d32')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      statusSheet.setFrozenRows(1);
    }

    // Find if email already exists
    const lastRow = statusSheet.getLastRow();
    let rowToUpdate = -1;

    if (lastRow > 1) {
      const emails = statusSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < emails.length; i++) {
        if (emails[i][0] === email) {
          rowToUpdate = i + 2; // Add 2 because of header row and 0-based index
          break;
        }
      }
    }

    const timestamp = new Date().toISOString();
    const updatedBy = 'Admin'; // Could be enhanced to track actual admin user

    if (rowToUpdate > 0) {
      // Update existing record
      statusSheet.getRange(rowToUpdate, 2, 1, 3).setValues([[newStatus, timestamp, updatedBy]]);
    } else {
      // Add new record
      statusSheet.appendRow([email, newStatus, timestamp, updatedBy]);
    }

    Logger.log(`Status updated for ${email} to ${newStatus}`);
    return {
      success: true,
      message: 'Status updated successfully'
    };

  } catch (error) {
    Logger.log('Error updating candidate status: ' + error.toString());
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * Sends email to a candidate
 * @param {string} candidateEmail - Recipient email
 * @param {string} subject - Email subject
 * @param {string} body - Email body
 * @returns {Object} Response object
 */
function sendEmailToCandidate(candidateEmail, subject, body) {
  try {
    MailApp.sendEmail(candidateEmail, subject, body);

    // Log email in EmailLog sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName('EmailLog');

    if (!logSheet) {
      logSheet = ss.insertSheet('EmailLog');
      logSheet.getRange(1, 1, 1, 5).setValues([['Timestamp', 'Recipient', 'Subject', 'Body', 'Sent By']]);
      logSheet.getRange(1, 1, 1, 5)
        .setBackground('#2e7d32')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      logSheet.setFrozenRows(1);
    }

    const timestamp = new Date().toISOString();
    logSheet.appendRow([timestamp, candidateEmail, subject, body, 'Admin']);

    Logger.log(`Email sent to ${candidateEmail}`);
    return {
      success: true,
      message: 'Email sent successfully'
    };

  } catch (error) {
    Logger.log('Error sending email: ' + error.toString());
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * Schedules an interview and sends invite email
 * @param {Object} interviewData - Interview details
 * @returns {Object} Response object
 */
function scheduleInterview(interviewData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Get or create Interviews sheet
    let interviewSheet = ss.getSheetByName('Interviews');
    if (!interviewSheet) {
      interviewSheet = ss.insertSheet('Interviews');
      interviewSheet.getRange(1, 1, 1, 7).setValues([['Scheduled At', 'Candidate Email', 'Interview Date (ISO)', 'Location', 'Interviewer', 'Notes', 'Status']]);
      interviewSheet.getRange(1, 1, 1, 7)
        .setBackground('#2e7d32')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      interviewSheet.setFrozenRows(1);
    }

    const timestamp = new Date().toISOString();

    // Extract time in 12-hour format from ISO date
    const interviewTime = formatTimeFromISO(interviewData.dateTime);

    // Format the interview date for display
    const interviewDateDisplay = interviewData.dateTime ?
      formatISODate(interviewData.dateTime) : '';

    interviewSheet.appendRow([
      timestamp,
      interviewData.candidateEmail,
      interviewData.dateTime, // ISO format: 2025-10-09T06:02:00.000Z
      interviewData.location,
      interviewData.interviewer,
      interviewData.notes,
      'Scheduled'
    ]);

    // Send interview invitation email
    const subject = 'Interview Invitation';
    const body = `Dear Candidate,

We are pleased to invite you for an interview.

Interview Details:
Date & Time: ${interviewDateDisplay} at ${interviewTime}
Location: ${interviewData.location}
Interviewer: ${interviewData.interviewer}

${interviewData.notes ? 'Additional Notes:\n' + interviewData.notes : ''}

Please confirm your availability for the scheduled time.

Best regards,
HR Team`;

    MailApp.sendEmail(interviewData.candidateEmail, subject, body);

    Logger.log(`Interview scheduled for ${interviewData.candidateEmail} on ${interviewData.dateTime}`);
    return {
      success: true,
      message: 'Interview scheduled and invitation sent'
    };

  } catch (error) {
    Logger.log('Error scheduling interview: ' + error.toString());
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * Reschedules an existing interview and sends update notification email
 * @param {Object} interviewData - Interview details including original scheduled time
 * @returns {Object} Response object
 */
function rescheduleInterview(interviewData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const interviewSheet = ss.getSheetByName('Interviews');

    if (!interviewSheet) {
      return {
        success: false,
        message: 'Interviews sheet not found'
      };
    }

    // Find the existing interview record
    const lastRow = interviewSheet.getLastRow();
    if (lastRow <= 1) {
      return {
        success: false,
        message: 'No interviews found to reschedule'
      };
    }

    const data = interviewSheet.getRange(2, 1, lastRow - 1, 7).getValues();
    let rowIndex = -1;

    // Find the interview by candidateEmail and originalScheduledAt
    for (let i = 0; i < data.length; i++) {
      if (data[i][1] === interviewData.candidateEmail &&
          data[i][0] === interviewData.originalScheduledAt) {
        rowIndex = i + 2; // Add 2 because of header row and 0-based index
        break;
      }
    }

    if (rowIndex === -1) {
      return {
        success: false,
        message: 'Interview not found'
      };
    }

    const timestamp = new Date().toISOString();

    // Extract time in 12-hour format from ISO date
    const interviewTime = formatTimeFromISO(interviewData.dateTime);

    // Format the interview date for display
    const interviewDateDisplay = interviewData.dateTime ?
      formatISODate(interviewData.dateTime) : '';

    // Update the existing interview record
    interviewSheet.getRange(rowIndex, 1, 1, 7).setValues([[
      timestamp, // Updated scheduled at timestamp
      interviewData.candidateEmail,
      interviewData.dateTime, // ISO format: 2025-10-09T06:02:00.000Z
      interviewData.location,
      interviewData.interviewer,
      interviewData.notes,
      'Rescheduled'
    ]]);

    // Send rescheduled interview notification email
    const subject = 'Interview Rescheduled - Updated Details';
    const body = `Dear Candidate,

Your interview has been rescheduled. Please find the updated details below:

Updated Interview Details:
Date & Time: ${interviewDateDisplay} at ${interviewTime}
Location: ${interviewData.location}
Interviewer: ${interviewData.interviewer}

${interviewData.notes ? 'Additional Notes:\n' + interviewData.notes : ''}

We apologize for any inconvenience. Please confirm your availability for the rescheduled time.

Best regards,
HR Team`;

    MailApp.sendEmail(interviewData.candidateEmail, subject, body);

    Logger.log(`Interview rescheduled for ${interviewData.candidateEmail} to ${interviewData.dateTime}`);
    return {
      success: true,
      message: 'Interview rescheduled and notification sent'
    };

  } catch (error) {
    Logger.log('Error rescheduling interview: ' + error.toString());
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * Formats time from ISO date string to 12-hour AM/PM format
 * @param {string} isoString - ISO 8601 date string
 * @returns {string} Formatted time string (e.g., "10:30:00 AM")
 */
function formatTimeFromISO(isoString) {
  try {
    const date = new Date(isoString);
    const hours = date.getHours();
    const minutes = String(date.getMinutes()).padStart(2, '0');
    const seconds = String(date.getSeconds()).padStart(2, '0');
    const ampm = hours >= 12 ? 'PM' : 'AM';
    const displayHour = hours % 12 || 12;

    return `${displayHour}:${minutes}:${seconds} ${ampm}`;
  } catch (error) {
    return '';
  }
}

/**
 * Gets the URL of a CV file from Google Drive
 * @param {string} fileName - Name of the CV file
 * @returns {string} File URL or empty string if not found
 */
function getCVUrl(fileName) {
  try {
    if (!fileName) return '';

    const folder = getCVFolder();
    const files = folder.getFilesByName(fileName);

    if (files.hasNext()) {
      const file = files.next();
      // Return the shareable link
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return file.getUrl();
    }

    // If exact name not found, try searching for files containing the email
    const searchQuery = fileName.split('_')[1] || fileName; // Try to extract email part
    const searchFiles = folder.searchFiles(`title contains '${searchQuery}'`);

    if (searchFiles.hasNext()) {
      const file = searchFiles.next();
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return file.getUrl();
    }

    return '';

  } catch (error) {
    Logger.log('Error getting CV URL: ' + error.toString());
    return '';
  }
}

/**
 * Sends bulk emails to multiple candidates
 * @param {Array} emails - Array of email addresses
 * @param {string} subject - Email subject
 * @param {string} bodyTemplate - Email body template
 * @returns {Object} Response object
 */
function sendBulkEmails(emails, subject, bodyTemplate) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName('EmailLog');

    if (!logSheet) {
      logSheet = ss.insertSheet('EmailLog');
      logSheet.getRange(1, 1, 1, 5).setValues([['Timestamp', 'Recipient', 'Subject', 'Body', 'Sent By']]);
      logSheet.getRange(1, 1, 1, 5)
        .setBackground('#2e7d32')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      logSheet.setFrozenRows(1);
    }

    const timestamp = new Date().toISOString();
    let successCount = 0;
    let failedEmails = [];

    emails.forEach(email => {
      try {
        MailApp.sendEmail(email, subject, bodyTemplate);
        logSheet.appendRow([timestamp, email, subject, bodyTemplate, 'Admin - Bulk']);
        successCount++;
      } catch (err) {
        failedEmails.push(email);
        Logger.log(`Failed to send email to ${email}: ${err.toString()}`);
      }
    });

    return {
      success: true,
      message: `Emails sent: ${successCount}/${emails.length}`,
      failedEmails: failedEmails
    };

  } catch (error) {
    Logger.log('Error sending bulk emails: ' + error.toString());
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * Gets interview schedule for a candidate
 * @param {string} candidateEmail - Candidate's email
 * @returns {Object} Interview details or null
 */
function getInterviewDetails(candidateEmail) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const interviewSheet = ss.getSheetByName('Interviews');

    if (!interviewSheet) {
      return null;
    }

    const lastRow = interviewSheet.getLastRow();
    if (lastRow <= 1) return null;

    const data = interviewSheet.getRange(2, 1, lastRow - 1, 7).getValues();

    // Find the latest interview for this candidate
    for (let i = data.length - 1; i >= 0; i--) {
      if (data[i][1] === candidateEmail) {
        return {
          scheduledAt: data[i][0], // ISO timestamp when scheduled
          interviewDateTime: data[i][2], // ISO format: 2025-10-09T06:02:00.000Z
          interviewTime: formatTimeFromISO(data[i][2]), // Extract time from ISO: 10:30:00 AM
          location: data[i][3],
          interviewer: data[i][4],
          notes: data[i][5],
          status: data[i][6]
        };
      }
    }

    return null;

  } catch (error) {
    Logger.log('Error getting interview details: ' + error.toString());
    return null;
  }
}

/**
 * Gets all scheduled interviews
 * @returns {Array} Array of all interview objects
 */
function getAllInterviews() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const interviewSheet = ss.getSheetByName('Interviews');

    if (!interviewSheet) {
      Logger.log('Interviews sheet not found');
      return [];
    }

    const lastRow = interviewSheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('No interviews found');
      return [];
    }

    const data = interviewSheet.getRange(2, 1, lastRow - 1, 7).getValues();

    // Get all applications to match candidate names
    const applications = getAllApplications();
    const candidateMap = new Map();
    applications.forEach(app => {
      candidateMap.set(app.email, `${app.firstName} ${app.lastName}`);
    });

    // Map data to interview objects
    const interviews = data.map(row => ({
      scheduledAt: row[0], // ISO timestamp when scheduled
      candidateEmail: row[1],
      candidateName: candidateMap.get(row[1]) || 'Unknown',
      interviewDateTime: row[2], // ISO format: 2025-10-09T06:02:00.000Z
      interviewTime: formatTimeFromISO(row[2]), // Extract time from ISO: 10:30:00 AM
      location: row[3],
      interviewer: row[4],
      notes: row[5],
      status: row[6]
    }));

    // Sort by interview date (newest first)
    interviews.sort((a, b) => {
      const dateA = new Date(a.interviewDateTime);
      const dateB = new Date(b.interviewDateTime);
      return dateB - dateA;
    });

    Logger.log(`Retrieved ${interviews.length} interviews`);
    return interviews;

  } catch (error) {
    Logger.log('Error fetching interviews: ' + error.toString());
    return [];
  }
}

/**
 * Deletes an application from the Applications sheet
 * @param {string} candidateEmail - Email of the candidate to delete
 * @returns {Object} Response object
 */
function deleteApplication(candidateEmail) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const appSheet = ss.getSheetByName('Applications');

    if (!appSheet) {
      return {
        success: false,
        message: 'Applications sheet not found'
      };
    }

    // Find the row with this email
    const lastRow = appSheet.getLastRow();
    if (lastRow <= 1) {
      return {
        success: false,
        message: 'No applications found'
      };
    }

    const emailColumn = 5; // Email is in column E (5th column)
    const emails = appSheet.getRange(2, emailColumn, lastRow - 1, 1).getValues();

    let rowToDelete = -1;
    for (let i = 0; i < emails.length; i++) {
      if (emails[i][0] === candidateEmail) {
        rowToDelete = i + 2; // Add 2 because of header row and 0-based index
        break;
      }
    }

    if (rowToDelete === -1) {
      return {
        success: false,
        message: 'Application not found'
      };
    }

    // Delete the row
    appSheet.deleteRow(rowToDelete);

    // Also delete from CandidateStatus sheet if exists
    const statusSheet = ss.getSheetByName('CandidateStatus');
    if (statusSheet) {
      const statusLastRow = statusSheet.getLastRow();
      if (statusLastRow > 1) {
        const statusEmails = statusSheet.getRange(2, 1, statusLastRow - 1, 1).getValues();
        for (let i = 0; i < statusEmails.length; i++) {
          if (statusEmails[i][0] === candidateEmail) {
            statusSheet.deleteRow(i + 2);
            break;
          }
        }
      }
    }

    Logger.log(`Application deleted for ${candidateEmail}`);
    return {
      success: true,
      message: 'Application deleted successfully'
    };

  } catch (error) {
    Logger.log('Error deleting application: ' + error.toString());
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * Updates an application in the Applications sheet
 * @param {Object} updateData - Object containing email and fields to update
 * @returns {Object} Response object
 */
function updateApplication(updateData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const appSheet = ss.getSheetByName('Applications');

    if (!appSheet) {
      return {
        success: false,
        message: 'Applications sheet not found'
      };
    }

    // Find the row with this email
    const lastRow = appSheet.getLastRow();
    if (lastRow <= 1) {
      return {
        success: false,
        message: 'No applications found'
      };
    }

    const emailColumn = 5; // Email is in column E (5th column)
    const emails = appSheet.getRange(2, emailColumn, lastRow - 1, 1).getValues();

    let rowToUpdate = -1;
    for (let i = 0; i < emails.length; i++) {
      if (emails[i][0] === updateData.email) {
        rowToUpdate = i + 2; // Add 2 because of header row and 0-based index
        break;
      }
    }

    if (rowToUpdate === -1) {
      return {
        success: false,
        message: 'Application not found'
      };
    }

    // Update the fields
    // Column indices: Mobile=6, YearsOfExperience=32, LastOrganisation=33, LastDesignation=34
    if (updateData.mobile !== undefined) {
      appSheet.getRange(rowToUpdate, 6).setValue(updateData.mobile);
    }
    if (updateData.yearsOfExperience !== undefined) {
      appSheet.getRange(rowToUpdate, 32).setValue(updateData.yearsOfExperience);
    }
    if (updateData.lastOrganisation !== undefined) {
      appSheet.getRange(rowToUpdate, 33).setValue(updateData.lastOrganisation);
    }
    if (updateData.lastDesignation !== undefined) {
      appSheet.getRange(rowToUpdate, 34).setValue(updateData.lastDesignation);
    }

    // Update status in CandidateStatus sheet if status is provided
    if (updateData.status !== undefined) {
      const result = updateCandidateStatus(updateData.email, updateData.status);
      if (!result.success) {
        Logger.log('Warning: Failed to update candidate status: ' + result.message);
      }
    }

    Logger.log(`Application updated for ${updateData.email}`);
    return {
      success: true,
      message: 'Application updated successfully'
    };

  } catch (error) {
    Logger.log('Error updating application: ' + error.toString());
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * Sets up demo data with 20 generic test applicants
 * Run this function manually to populate the system with test data
 * All data is fictional and for testing purposes only
 */
function setupDemoData() {
  Logger.log('=== Setting up Demo Data ===');

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Applications');

    if (!sheet) {
      throw new Error('Applications sheet not found. Please run setupSystem() first.');
    }

    // Check if headers exist, if not add them
    const lastRow = sheet.getLastRow();
    if (lastRow === 0) {
      const headers = [
        'Timestamp', 'Applying For',
        'First Name', 'Last Name', 'Email', 'Mobile', 'DOB', 'Religion',
        'Marital Status', 'Gender', 'Nationality', 'Current Address', 'Permanent Address',
        'SSC Institute', 'SSC Board', 'SSC Major', 'SSC Year', 'SSC Grade',
        'HSC Institute', 'HSC Board', 'HSC Major', 'HSC Year', 'HSC Grade',
        'Bachelor Degree', 'Bachelor Institute', 'Bachelor Major', 'Bachelor Grade',
        'Postgraduate Degree', 'Postgraduate Institute', 'Postgraduate Major', 'Postgraduate Grade',
        'Years of Experience', 'Last Organisation', 'Last Designation',
        'Joining Date', 'End Date', 'Currently Working', 'CV URL'
      ];
      sheet.appendRow(headers);

      // Format header row
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground('#2e7d32')
        .setFontColor('#ffffff')
        .setFontWeight('bold');

      sheet.setFrozenRows(1);
      Logger.log('Headers added to Applications sheet');
    }

    // Generic demo data for testing
    const demoApplicants = [
      {
        applyingFor: 'Software Engineer',
        firstName: 'Demo', lastName: 'Employee1',
        email: 'demo.employee1@example.com', mobile: '+1-000-111-0001',
        dob: '1998-03-15T00:00:00.000Z', religion: 'Not Specified', maritalStatus: 'Single',
        gender: 'Male', nationality: 'Country1',
        currentAddress: 'Address Line 1, Street 1, City 1, State 1',
        permanentAddress: 'Address Line 1, Street 1, City 1, State 1',
        sscInstitute: 'High School 1', sscBoard: 'Board 1', sscMajor: 'Science', sscYear: '2014', sscGrade: 'A',
        hscInstitute: 'College 1', hscBoard: 'Board 1', hscMajor: 'Computer Science', hscYear: '2016', hscGrade: 'A',
        bachelorDegree: 'B.Sc. Computer Science', bachelorInstitute: 'University 1', bachelorMajor: 'Computer Science', bachelorGrade: '3.8/4.0',
        postgraduateDegree: 'M.Sc. Software Engineering', postgraduateInstitute: 'University 2', postgraduateMajor: 'Software Engineering', postgraduateGrade: '3.9/4.0',
        yearsOfExperience: '5', lastOrganisation: 'Company 1', lastDesignation: 'Software Engineer',
        joiningDate: '2019-01-15T00:00:00.000Z', endDate: '2024-08-30T00:00:00.000Z', currentlyWorking: 'No'
      },
      {
        applyingFor: 'HR Manager',
        firstName: 'Demo', lastName: 'Employee2',
        email: 'demo.employee2@example.com', mobile: '+1-000-111-0002',
        dob: '1995-07-22T00:00:00.000Z', religion: 'Not Specified', maritalStatus: 'Married',
        gender: 'Female', nationality: 'Country1',
        currentAddress: 'Address Line 2, Street 2, City 1, State 1',
        permanentAddress: 'Address Line 2, Street 2, City 1, State 1',
        sscInstitute: 'High School 2', sscBoard: 'Board', sscMajor: 'Science', sscYear: '2011', sscGrade: 'A',
        hscInstitute: 'High School 2', hscBoard: 'Board', hscMajor: 'Pre-Medical', hscYear: '2013', hscGrade: 'A',
        bachelorDegree: 'BBA', bachelorInstitute: 'University 2', bachelorMajor: 'Human Resource Management', bachelorGrade: '3.7/4.0',
        postgraduateDegree: 'MBA', postgraduateInstitute: 'University 2', postgraduateMajor: 'HR Management', postgraduateGrade: '3.8/4.0',
        yearsOfExperience: '7', lastOrganisation: 'Company Country1', lastDesignation: 'HR Manager',
        joiningDate: '2017-03-01T00:00:00.000Z', endDate: '', currentlyWorking: 'Yes'
      },
      {
        applyingFor: 'Graphic Designer',
        firstName: 'Demo', lastName: 'Employee3',
        email: 'demo.employee3@example.com', mobile: '+1-000-111-0003',
        dob: '1997-11-10T00:00:00.000Z', religion: 'Not Specified', maritalStatus: 'Single',
        gender: 'Male', nationality: 'Country1',
        currentAddress: 'Address Line 3, Street 3, City 3, State 3',
        permanentAddress: 'Address Line 3B, Street 3B, City 6, State 6',
        sscInstitute: 'High School', sscBoard: 'Board', sscMajor: 'Arts', sscYear: '2013', sscGrade: 'A',
        hscInstitute: 'College 3', hscBoard: 'Board', hscMajor: 'Fine Arts', hscYear: '2015', hscGrade: 'A',
        bachelorDegree: 'B.Design', bachelorInstitute: 'University 3', bachelorMajor: 'Graphic Design', bachelorGrade: '3.6/4.0',
        postgraduateDegree: '', postgraduateInstitute: '', postgraduateMajor: '', postgraduateGrade: '',
        yearsOfExperience: '4', lastOrganisation: 'Company', lastDesignation: 'Senior Graphic Designer',
        joiningDate: '2020-06-01T00:00:00.000Z', endDate: '2024-09-15T00:00:00.000Z', currentlyWorking: 'No'
      },
      {
        applyingFor: 'Marketing Manager',
        firstName: 'Demo', lastName: 'Employee4',
        email: 'demo.employee4@example.com', mobile: '+1-000-111-0004',
        dob: '1994-05-18T00:00:00.000Z', religion: 'Not Specified', maritalStatus: 'Married',
        gender: 'Female', nationality: 'Country1',
        currentAddress: 'Address Line 4, Street 4, City 2, State 2',
        permanentAddress: 'Address Line 4, Street 4, City 2, State 2',
        sscInstitute: 'High School 4', sscBoard: 'Board', sscMajor: 'Science', sscYear: '2010', sscGrade: 'A+',
        hscInstitute: 'High School 4', hscBoard: 'Board', hscMajor: 'Pre-Engineering', hscYear: '2012', hscGrade: 'A',
        bachelorDegree: 'BBA', bachelorInstitute: 'University', bachelorMajor: 'Marketing', bachelorGrade: '3.9/4.0',
        postgraduateDegree: 'MBA', postgraduateInstitute: 'University', postgraduateMajor: 'Marketing', postgraduateGrade: '3.9/4.0',
        yearsOfExperience: '8', lastOrganisation: 'Company Country1', lastDesignation: 'Marketing Manager',
        joiningDate: '2016-01-10T00:00:00.000Z', endDate: '', currentlyWorking: 'Yes'
      },
      {
        applyingFor: 'Data Analyst',
        firstName: 'Demo', lastName: 'Employee5',
        email: 'demo.employee5@example.com', mobile: '+1-000-111-0005',
        dob: '1999-02-28T00:00:00.000Z', religion: 'Not Specified', maritalStatus: 'Single',
        gender: 'Male', nationality: 'Country1',
        currentAddress: 'Address Line 5, Street 5, City 4, State 4',
        permanentAddress: 'Address Line 5B, Street 5B, City 4, State 4',
        sscInstitute: 'High School 5', sscBoard: 'Board', sscMajor: 'Science', sscYear: '2015', sscGrade: 'A',
        hscInstitute: 'High School 5', hscBoard: 'Board', hscMajor: 'Pre-Engineering', hscYear: '2017', hscGrade: 'A+',
        bachelorDegree: 'B.Sc. Data Science', bachelorInstitute: 'University 5', bachelorMajor: 'Data Science', bachelorGrade: '3.7/4.0',
        postgraduateDegree: '', postgraduateInstitute: '', postgraduateMajor: '', postgraduateGrade: '',
        yearsOfExperience: '3', lastOrganisation: 'Company', lastDesignation: 'Data Analyst',
        joiningDate: '2021-07-01T00:00:00.000Z', endDate: '', currentlyWorking: 'Yes'
      },
      {
        applyingFor: 'Sales Executive',
        firstName: 'Demo', lastName: 'Employee6',
        email: 'demo.employee6@example.com', mobile: '+1-000-111-0006',
        dob: '1996-09-12T00:00:00.000Z', religion: 'Not Specified', maritalStatus: 'Single',
        gender: 'Female', nationality: 'Country1',
        currentAddress: 'Address Line 6, Street 6, City 2, State 2',
        permanentAddress: 'Address Line 6, Street 6, City 2, State 2',
        sscInstitute: 'High School 6', sscBoard: 'Board', sscMajor: 'Science', sscYear: '2012', sscGrade: 'A',
        hscInstitute: 'College 6', hscBoard: 'Board', hscMajor: 'Commerce', hscYear: '2014', hscGrade: 'A',
        bachelorDegree: 'BBA', bachelorInstitute: 'University', bachelorMajor: 'Marketing', bachelorGrade: '3.5/4.0',
        postgraduateDegree: '', postgraduateInstitute: '', postgraduateMajor: '', postgraduateGrade: '',
        yearsOfExperience: '6', lastOrganisation: 'Company', lastDesignation: 'Senior Sales Executive',
        joiningDate: '2018-04-15T00:00:00.000Z', endDate: '2024-07-20T00:00:00.000Z', currentlyWorking: 'No'
      },
      {
        applyingFor: 'Network Engineer',
        firstName: 'Demo', lastName: 'Employee7',
        email: 'demo.employee7@example.com', mobile: '+1-000-111-0007',
        dob: '1997-12-05T00:00:00.000Z', religion: 'Not Specified', maritalStatus: 'Single',
        gender: 'Male', nationality: 'Country1',
        currentAddress: 'Address Line 7, Street 7, City 1, State 1',
        permanentAddress: 'Address Line 7, Street 7, City 1, State 1',
        sscInstitute: 'High School 7', sscBoard: 'Board', sscMajor: 'Science', sscYear: '2013', sscGrade: 'A+',
        hscInstitute: 'College 7', hscBoard: 'Board', hscMajor: 'Pre-Engineering', hscYear: '2015', sscGrade: 'A+',
        bachelorDegree: 'B.E. Telecommunications', bachelorInstitute: 'University 7', bachelorMajor: 'Telecommunications', bachelorGrade: '3.6/4.0',
        postgraduateDegree: '', postgraduateInstitute: '', postgraduateMajor: '', postgraduateGrade: '',
        yearsOfExperience: '4', lastOrganisation: 'Company', lastDesignation: 'Network Engineer',
        joiningDate: '2020-02-01T00:00:00.000Z', endDate: '', currentlyWorking: 'Yes'
      },
      {
        applyingFor: 'Finance Manager',
        firstName: 'Demo', lastName: 'Employee8',
        email: 'demo.employee8@example.com', mobile: '+1-000-111-0008',
        dob: '1993-04-08T00:00:00.000Z', religion: 'Not Specified', maritalStatus: 'Married',
        gender: 'Female', nationality: 'Country1',
        currentAddress: 'Address Line 8, Street 8, City 2, State 2',
        permanentAddress: 'Address Line 8, Street 8, City 2, State 2',
        sscInstitute: 'High School 8', sscBoard: 'Board', sscMajor: 'Science', sscYear: '2009', sscGrade: 'A+',
        hscInstitute: 'High School 8', hscBoard: 'Board', hscMajor: 'Commerce', hscYear: '2011', sscGrade: 'A+',
        bachelorDegree: 'B.Com', bachelorInstitute: 'University 8', bachelorMajor: 'Accounting & Finance', bachelorGrade: '3.8/4.0',
        postgraduateDegree: 'MBA', postgraduateInstitute: 'University', postgraduateMajor: 'Finance', postgraduateGrade: '3.9/4.0',
        yearsOfExperience: '9', lastOrganisation: 'Company', lastDesignation: 'Finance Manager',
        joiningDate: '2015-03-01T00:00:00.000Z', endDate: '', currentlyWorking: 'Yes'
      },
      {
        applyingFor: 'Content Writer',
        firstName: 'Demo', lastName: 'Employee9',
        email: 'demo.employee9@example.com', mobile: '+1-000-111-0009',
        dob: '1998-08-20T00:00:00.000Z', religion: 'Not Specified', maritalStatus: 'Single',
        gender: 'Male', nationality: 'Country1',
        currentAddress: 'Address Line 9, Street 9, City 3, State 3',
        permanentAddress: 'Address Line 9, Street 9, City 3, State 3',
        sscInstitute: 'High School 9', sscBoard: 'Board', sscMajor: 'Arts', sscYear: '2014', sscGrade: 'A',
        hscInstitute: 'High School 9', hscBoard: 'Board', hscMajor: 'Humanities', hscYear: '2016', hscGrade: 'A',
        bachelorDegree: 'BA Mass Communication', bachelorInstitute: 'University', bachelorMajor: 'Journalism', bachelorGrade: '3.7/4.0',
        postgraduateDegree: '', postgraduateInstitute: '', postgraduateMajor: '', postgraduateGrade: '',
        yearsOfExperience: '4', lastOrganisation: 'Company', lastDesignation: 'Senior Content Writer',
        joiningDate: '2020-09-01T00:00:00.000Z', endDate: '2024-08-31T00:00:00.000Z', currentlyWorking: 'No'
      },
      {
        applyingFor: 'UI/UX Designer',
        firstName: 'Demo', lastName: 'Employee10',
        email: 'demo.employee10@example.com', mobile: '+1-000-111-0010',
        dob: '1996-01-14T00:00:00.000Z', religion: 'Not Specified', maritalStatus: 'Single',
        gender: 'Female', nationality: 'Country1',
        currentAddress: 'Address Line 10, Street 10, City 1, State 1',
        permanentAddress: 'Address Line 10, Street 10, City 1, State 1',
        sscInstitute: 'High School 10', sscBoard: 'Board', sscMajor: 'Arts', sscYear: '2012', sscGrade: 'A*',
        hscInstitute: 'High School 10', hscBoard: 'Board', hscMajor: 'Arts & Design', hscYear: '2014', hscGrade: 'A*',
        bachelorDegree: 'B.Design', bachelorInstitute: 'University 10', bachelorMajor: 'Communication Design', bachelorGrade: '3.8/4.0',
        postgraduateDegree: '', postgraduateInstitute: '', postgraduateMajor: '', postgraduateGrade: '',
        yearsOfExperience: '5', lastOrganisation: 'Company', lastDesignation: 'UI/UX Designer',
        joiningDate: '2019-05-01T00:00:00.000Z', endDate: '', currentlyWorking: 'Yes'
      },
      {
        applyingFor: 'Project Manager',
        firstName: 'Demo', lastName: 'Employee11',
        email: 'demo.employee11@example.com', mobile: '+1-000-111-0011',
        dob: '1992-06-25T00:00:00.000Z', religion: 'Not Specified', maritalStatus: 'Married',
        gender: 'Male', nationality: 'Country1',
        currentAddress: 'Address Line 11, Street 11, City 2, State 2',
        permanentAddress: 'Address Line 11, Street 11, City 2, State 2',
        sscInstitute: 'High School 11', sscBoard: 'Board', sscMajor: 'Science', sscYear: '2008', sscGrade: 'A*',
        hscInstitute: 'High School 11', hscBoard: 'Board', hscMajor: 'Mathematics', hscYear: '2010', hscGrade: 'A*',
        bachelorDegree: 'B.Sc. Civil Engineering', bachelorInstitute: 'University', bachelorMajor: 'Civil Engineering', bachelorGrade: '3.7/4.0',
        postgraduateDegree: 'MBA', postgraduateInstitute: 'University', postgraduateMajor: 'Project Management', postgraduateGrade: '3.8/4.0',
        yearsOfExperience: '10', lastOrganisation: 'Company', lastDesignation: 'Project Manager',
        joiningDate: '2014-01-15T00:00:00.000Z', endDate: '', currentlyWorking: 'Yes'
      },
      {
        applyingFor: 'Business Analyst',
        firstName: 'Demo', lastName: 'Employee12',
        email: 'demo.employee12@example.com', mobile: '+1-000-111-0012',
        dob: '1995-10-03T00:00:00.000Z', religion: 'Not Specified', maritalStatus: 'Single',
        gender: 'Female', nationality: 'Country1',
        currentAddress: 'Address Line 12, Street 12, City 4, State 4',
        permanentAddress: 'Address Line 12B, Street 12B, City 4, State 4',
        sscInstitute: 'High School 12', sscBoard: 'Board', sscMajor: 'Science', sscYear: '2011', sscGrade: 'A+',
        hscInstitute: 'High School 12', hscBoard: 'Board', hscMajor: 'Pre-Engineering', hscYear: '2013', hscGrade: 'A',
        bachelorDegree: 'B.Sc. Business Administration', bachelorInstitute: 'University 12', bachelorMajor: 'Business Administration', bachelorGrade: '3.6/4.0',
        postgraduateDegree: 'MS Business Analytics', postgraduateInstitute: 'University', postgraduateMajor: 'Business Analytics', postgraduateGrade: '3.8/4.0',
        yearsOfExperience: '6', lastOrganisation: 'Company', lastDesignation: 'Business Analyst',
        joiningDate: '2018-07-01T00:00:00.000Z', endDate: '', currentlyWorking: 'Yes'
      },
      {
        applyingFor: 'DevOps Engineer',
        firstName: 'Demo', lastName: 'Employee13',
        email: 'demo.employee13@example.com', mobile: '+1-000-111-0013',
        dob: '1997-03-17T00:00:00.000Z', religion: 'Not Specified', maritalStatus: 'Single',
        gender: 'Male', nationality: 'Country1',
        currentAddress: 'Address Line 13, Street 13, City 1, State 1',
        permanentAddress: 'Address Line 13, Street 13, City 1, State 1',
        sscInstitute: 'High School 2', sscBoard: 'Board', sscMajor: 'Science', sscYear: '2013', sscGrade: 'A*',
        hscInstitute: 'High School 2', hscBoard: 'Board', hscMajor: 'Computer Science', hscYear: '2015', hscGrade: 'A*',
        bachelorDegree: 'B.Sc. Computer Science', bachelorInstitute: 'University 5', bachelorMajor: 'Computer Science', bachelorGrade: '3.9/4.0',
        postgraduateDegree: '', postgraduateInstitute: '', postgraduateMajor: '', postgraduateGrade: '',
        yearsOfExperience: '5', lastOrganisation: 'Company', lastDesignation: 'DevOps Engineer',
        joiningDate: '2019-08-01T00:00:00.000Z', endDate: '2024-09-30T00:00:00.000Z', currentlyWorking: 'No'
      },
      {
        applyingFor: 'Accountant',
        firstName: 'Demo', lastName: 'Employee14',
        email: 'demo.employee14@example.com', mobile: '+1-000-111-0014',
        dob: '1994-11-29T00:00:00.000Z', religion: 'Not Specified', maritalStatus: 'Married',
        gender: 'Female', nationality: 'Country1',
        currentAddress: 'Address Line 14, Street 14, City 1, State 1',
        permanentAddress: 'Address Line 14, Street 14, City 1, State 1',
        sscInstitute: 'High School 14', sscBoard: 'Board', sscMajor: 'Science', sscYear: '2010', sscGrade: 'A',
        hscInstitute: 'High School 14', hscBoard: 'Board', hscMajor: 'Commerce', hscYear: '2012', hscGrade: 'A+',
        bachelorDegree: 'B.Com', bachelorInstitute: 'University', bachelorMajor: 'Accounting', bachelorGrade: '3.5/4.0',
        postgraduateDegree: 'University 14', postgraduateInstitute: 'University 14', postgraduateMajor: 'Accounting & Finance', postgraduateGrade: 'Pass',
        yearsOfExperience: '8', lastOrganisation: 'Company', lastDesignation: 'Senior Accountant',
        joiningDate: '2016-05-01T00:00:00.000Z', endDate: '', currentlyWorking: 'Yes'
      },
      {
        applyingFor: 'Quality Assurance Engineer',
        firstName: 'Demo', lastName: 'Employee15',
        email: 'demo.employee15@example.com', mobile: '+1-000-111-0015',
        dob: '1998-07-11T00:00:00.000Z', religion: 'Not Specified', maritalStatus: 'Single',
        gender: 'Male', nationality: 'Country1',
        currentAddress: 'Address Line 15, Street 15, City 2, State 2',
        permanentAddress: 'Address Line 15B, Street 15B, City 8, State 8',
        sscInstitute: 'High School', sscBoard: 'Board', sscMajor: 'Science', sscYear: '2014', sscGrade: 'A',
        hscInstitute: 'College 15', hscBoard: 'Board', hscMajor: 'ICS', hscYear: '2016', hscGrade: 'A',
        bachelorDegree: 'B.Sc. Software Engineering', bachelorInstitute: 'University', bachelorMajor: 'Software Engineering', bachelorGrade: '3.4/4.0',
        postgraduateDegree: '', postgraduateInstitute: '', postgraduateMajor: '', postgraduateGrade: '',
        yearsOfExperience: '4', lastOrganisation: 'Company', lastDesignation: 'QA Engineer',
        joiningDate: '2020-10-01T00:00:00.000Z', endDate: '', currentlyWorking: 'Yes'
      },
      {
        applyingFor: 'Digital Marketing Specialist',
        firstName: 'Demo', lastName: 'Employee16',
        email: 'demo.employee16@example.com', mobile: '+1-000-111-0016',
        dob: '1996-02-18T00:00:00.000Z', religion: 'Not Specified', maritalStatus: 'Single',
        gender: 'Female', nationality: 'Country1',
        currentAddress: 'Address Line 16, Street 16, City 3, State 3',
        permanentAddress: 'Address Line 16, Street 16, City 3, State 3',
        sscInstitute: 'High School 4', sscBoard: 'Board', sscMajor: 'Arts', sscYear: '2012', sscGrade: 'A',
        hscInstitute: 'High School 4', hscBoard: 'Board', hscMajor: 'Humanities', hscYear: '2014', hscGrade: 'A',
        bachelorDegree: 'BBA', bachelorInstitute: 'University 16', bachelorMajor: 'Marketing', bachelorGrade: '3.6/4.0',
        postgraduateDegree: '', postgraduateInstitute: '', postgraduateMajor: '', postgraduateGrade: '',
        yearsOfExperience: '5', lastOrganisation: 'Company', lastDesignation: 'Digital Marketing Specialist',
        joiningDate: '2019-11-01T00:00:00.000Z', endDate: '', currentlyWorking: 'Yes'
      },
      {
        applyingFor: 'Mechanical Engineer',
        firstName: 'Demo', lastName: 'Employee17',
        email: 'demo.employee17@example.com', mobile: '+1-000-111-0017',
        dob: '1995-09-08T00:00:00.000Z', religion: 'Not Specified', maritalStatus: 'Married',
        gender: 'Male', nationality: 'Country1',
        currentAddress: 'Address Line 17, Street 17, City 2, State 2',
        permanentAddress: 'Address Line 17B, Street 17B, City 7, State 7',
        sscInstitute: 'High School', sscBoard: 'Board', sscMajor: 'Science', sscYear: '2011', sscGrade: 'A',
        hscInstitute: 'College', hscBoard: 'Board', hscMajor: 'Pre-Engineering', hscYear: '2013', hscGrade: 'A+',
        bachelorDegree: 'B.E. Mechanical Engineering', bachelorInstitute: 'University', bachelorMajor: 'Mechanical Engineering', bachelorGrade: '3.5/4.0',
        postgraduateDegree: 'MS Engineering Management', postgraduateInstitute: 'University', postgraduateMajor: 'Engineering Management', postgraduateGrade: '3.7/4.0',
        yearsOfExperience: '7', lastOrganisation: 'Company', lastDesignation: 'Senior Mechanical Engineer',
        joiningDate: '2017-02-01T00:00:00.000Z', endDate: '', currentlyWorking: 'Yes'
      },
      {
        applyingFor: 'Customer Support Representative',
        firstName: 'Demo', lastName: 'Employee18',
        email: 'demo.employee18@example.com', mobile: '+1-000-111-0018',
        dob: '1999-04-22T00:00:00.000Z', religion: 'Not Specified', maritalStatus: 'Single',
        gender: 'Female', nationality: 'Country1',
        currentAddress: 'Address Line 18, Street 18, City 1, State 1',
        permanentAddress: 'Address Line 18, Street 18, City 1, State 1',
        sscInstitute: 'High School 18', sscBoard: 'Board', sscMajor: 'Science', sscYear: '2015', sscGrade: 'A',
        hscInstitute: 'College 18', hscBoard: 'Board', hscMajor: 'Pre-Medical', hscYear: '2017', hscGrade: 'A',
        bachelorDegree: 'BA English', bachelorInstitute: 'University', bachelorMajor: 'English Literature', bachelorGrade: '3.3/4.0',
        postgraduateDegree: '', postgraduateInstitute: '', postgraduateMajor: '', postgraduateGrade: '',
        yearsOfExperience: '3', lastOrganisation: 'Company', lastDesignation: 'Customer Support Representative',
        joiningDate: '2021-03-15T00:00:00.000Z', endDate: '', currentlyWorking: 'Yes'
      },
      {
        applyingFor: 'Civil Engineer',
        firstName: 'Demo', lastName: 'Employee19',
        email: 'demo.employee19@example.com', mobile: '+1-000-111-0019',
        dob: '1994-12-30T00:00:00.000Z', religion: 'Not Specified', maritalStatus: 'Married',
        gender: 'Male', nationality: 'Country1',
        currentAddress: 'Address Line 19, Street 19, City 5, State 5',
        permanentAddress: 'Address Line 19, Street 19, City 5, State 5',
        sscInstitute: 'High School 19', sscBoard: 'Board', sscMajor: 'Science', sscYear: '2010', sscGrade: 'A',
        hscInstitute: 'High School 19', hscBoard: 'Board', hscMajor: 'Pre-Engineering', hscYear: '2012', hscGrade: 'A+',
        bachelorDegree: 'B.E. Civil Engineering', bachelorInstitute: 'University 19', bachelorMajor: 'Civil Engineering', bachelorGrade: '3.6/4.0',
        postgraduateDegree: '', postgraduateInstitute: '', postgraduateMajor: '', postgraduateGrade: '',
        yearsOfExperience: '8', lastOrganisation: 'Company', lastDesignation: 'Civil Engineer',
        joiningDate: '2016-06-01T00:00:00.000Z', endDate: '', currentlyWorking: 'Yes'
      },
      {
        applyingFor: 'Operations Manager',
        firstName: 'Demo', lastName: 'Employee20',
        email: 'demo.employee20@example.com', mobile: '+1-000-111-0020',
        dob: '1993-05-27T00:00:00.000Z', religion: 'Not Specified', maritalStatus: 'Married',
        gender: 'Female', nationality: 'Country1',
        currentAddress: 'Address Line 20, Street 20, City 1, State 1',
        permanentAddress: 'Address Line 20, Street 20, City 1, State 1',
        sscInstitute: 'High School 20', sscBoard: 'Board', sscMajor: 'Science', sscYear: '2009', sscGrade: 'A+',
        hscInstitute: 'College 20', hscBoard: 'Board', hscMajor: 'Commerce', hscYear: '2011', hscGrade: 'A+',
        bachelorDegree: 'BBA', bachelorInstitute: 'University 2', bachelorMajor: 'Supply Chain Management', bachelorGrade: '3.8/4.0',
        postgraduateDegree: 'MBA', postgraduateInstitute: 'University 2', postgraduateMajor: 'Operations Management', postgraduateGrade: '3.9/4.0',
        yearsOfExperience: '9', lastOrganisation: 'Company Country1', lastDesignation: 'Operations Manager',
        joiningDate: '2015-04-01T00:00:00.000Z', endDate: '', currentlyWorking: 'Yes'
      }
    ];
    
    let added = 0;
    demoApplicants.forEach((applicant, index) => {
      const timestamp = new Date(Date.now() - (20 - index) * 24 * 60 * 60 * 1000).toISOString(); // Stagger by days
      const row = [
        timestamp,
        applicant.applyingFor,
        applicant.firstName,
        applicant.lastName,
        applicant.email,
        applicant.mobile,
        applicant.dob,
        applicant.religion,
        applicant.maritalStatus,
        applicant.gender,
        applicant.nationality,
        applicant.currentAddress,
        applicant.permanentAddress,
        applicant.sscInstitute,
        applicant.sscBoard,
        applicant.sscMajor,
        applicant.sscYear,
        applicant.sscGrade,
        applicant.hscInstitute,
        applicant.hscBoard,
        applicant.hscMajor,
        applicant.hscYear,
        applicant.hscGrade,
        applicant.bachelorDegree,
        applicant.bachelorInstitute,
        applicant.bachelorMajor,
        applicant.bachelorGrade,
        applicant.postgraduateDegree,
        applicant.postgraduateInstitute,
        applicant.postgraduateMajor,
        applicant.postgraduateGrade,
        applicant.yearsOfExperience,
        applicant.lastOrganisation,
        applicant.lastDesignation,
        applicant.joiningDate,
        applicant.endDate,
        applicant.currentlyWorking,
        `${applicant.firstName}_${applicant.lastName}_CV.pdf`
      ];
      sheet.appendRow(row);
      added++;
    });
    
    Logger.log(`=== Demo Data Setup Complete ===`);
    Logger.log(`Added ${added} generic demo applicants to the system`);
    
    return {
      success: true,
      message: `Successfully added ${added} demo applicants`,
      count: added
    };
    
  } catch (error) {
    Logger.log('Demo data setup error: ' + error.toString());
    return {
      success: false,
      message: 'Demo data setup failed: ' + error.message
    };
  }
}
/**
 * Saves email templates to EmailTemplates sheet
 * @param {Object} templates - Email templates object
 * @returns {Object} Response object
 */
function saveEmailTemplates(templates) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Get or create EmailTemplates sheet
    let templatesSheet = ss.getSheetByName('EmailTemplates');
    if (!templatesSheet) {
      templatesSheet = ss.insertSheet('EmailTemplates');
      templatesSheet.getRange(1, 1, 1, 2).setValues([['Template Name', 'Template Content']]);
      templatesSheet.getRange(1, 1, 1, 2)
        .setBackground('#2e7d32')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      templatesSheet.setFrozenRows(1);
    }

    // Clear existing data (keep header)
    const lastRow = templatesSheet.getLastRow();
    if (lastRow > 1) {
      templatesSheet.getRange(2, 1, lastRow - 1, 2).clear();
    }

    // Save templates
    const templateData = [
      ['confirmation', templates.confirmation],
      ['shortlist', templates.shortlist],
      ['interview', templates.interview],
      ['rejection', templates.rejection]
    ];

    templatesSheet.getRange(2, 1, templateData.length, 2).setValues(templateData);

    Logger.log('Email templates saved successfully');
    return {
      success: true,
      message: 'Email templates saved successfully'
    };

  } catch (error) {
    Logger.log('Error saving email templates: ' + error.toString());
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * Gets email templates from EmailTemplates sheet
 * @returns {Object} Templates object or default templates
 */
function getEmailTemplates() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const templatesSheet = ss.getSheetByName('EmailTemplates');

    // Return defaults if sheet doesn't exist
    if (!templatesSheet) {
      return getDefaultEmailTemplates();
    }

    const lastRow = templatesSheet.getLastRow();
    if (lastRow <= 1) {
      return getDefaultEmailTemplates();
    }

    // Read templates from sheet
    const data = templatesSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    const templates = {};

    data.forEach(row => {
      if (row[0]) {
        templates[row[0]] = row[1];
      }
    });

    // Return saved templates or defaults
    return {
      confirmation: templates.confirmation || getDefaultEmailTemplates().confirmation,
      shortlist: templates.shortlist || getDefaultEmailTemplates().shortlist,
      interview: templates.interview || getDefaultEmailTemplates().interview,
      rejection: templates.rejection || getDefaultEmailTemplates().rejection
    };

  } catch (error) {
    Logger.log('Error getting email templates: ' + error.toString());
    return getDefaultEmailTemplates();
  }
}

/**
 * Returns default email templates
 * @returns {Object} Default templates
 */
function getDefaultEmailTemplates() {
  return {
    confirmation: `Dear [Candidate Name],

Thank you for your interest in our position. We have received your application and our team is currently reviewing it.

We will contact you within the next few days regarding the next steps in the selection process.

Best regards,
HR Team`,

    shortlist: `Dear [Candidate Name],

Congratulations! We are pleased to inform you that you have been shortlisted for the [Position] role.

We would like to invite you for an interview. Our HR team will contact you shortly to schedule a convenient time.

Best regards,
HR Team`,

    interview: `Dear [Candidate Name],

We are pleased to invite you for an interview for the [Position] role.

Interview Details:
Date: [Date]
Time: [Time]
Location: [Location/Meeting Link]
Interviewer: [Interviewer Name]

Please confirm your availability for the scheduled time.

Best regards,
HR Team`,

    rejection: `Dear [Candidate Name],

Thank you for your interest in the [Position] role and for taking the time to apply.

After careful consideration, we have decided to move forward with other candidates whose qualifications more closely match our current requirements.

We will keep your resume on file for future opportunities that match your profile.

Best regards,
HR Team`
  };
}
