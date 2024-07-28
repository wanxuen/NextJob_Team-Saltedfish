function doGet() {
  // var page = e.parameter.page || 'Index'; // Default to 'Index' if no page parameter
  // console.log(page);
  // console.log("run")
  // return HtmlService.createHtmlOutputFromFile(page);
  return HtmlService.createHtmlOutputFromFile('Index');

}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getCandidates() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Candidates');
  const data = sheet.getDataRange().getValues();

  const candidates = data.slice(1).map(row => {
    try {
      // Extract file ID from the provided link format
      const url = row[0]; // Assuming the URL is in the first column
      const fileIdMatch = url.match(/\/d\/([^\/]*)\/view/);
      if (fileIdMatch) {
        const fileId = fileIdMatch[1];
        return {
          profilePic: `https://drive.google.com/uc?export=view&id=${fileId}`,
          name: row[1],
          email: row[2],
          resume: row[3],
          status: row[4],
          interviewSession: row[5]
        };
      } else {
        // Handle invalid URLs more gracefully (optional)
        console.warn(`Invalid file URL format in row: ${row[0]}`);
        return null; // or provide a default profile picture URL
      }
    } catch (error) {
      Logger.log(`Error processing row: ${error.message}`);
      return null; // or handle the error as needed
    }
  }).filter(candidate => candidate !== null);

  return candidates;
}

function selectCandidate(email) {
  updateCandidateStatus(email, 'Selected');
}

function declineCandidate(email) {
  updateCandidateStatus(email, 'Declined');
}

function updateCandidateStatus(email, status) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Candidates');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === email) {
      sheet.getRange(i + 1, 5).setValue(status);
      if (status === 'Selected') {
        const selectedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('selectedCandidates');
        selectedSheet.appendRow(data[i]);
      }
      break;
    }
  }
      resumeAnalysis()
      generateCharts();

}

function getPieChartData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Candidates');
  const data = sheet.getDataRange().getValues();
  const statusCount = { Pending: 0, Scheduled: 0, Selected: 0, Declined: 0 };
  for (let i = 1; i < data.length; i++) {
    statusCount[data[i][4]]++;
  }
  return [
    ['Status', 'Count'],
    ['Pending', statusCount.Pending],
    ['Scheduled', statusCount.Scheduled],
    ['Selected', statusCount.Selected],
    ['Declined', statusCount.Declined]
  ];
}


//select candidate
function getSelectedCandidates() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('selectedCandidates');
  const data = sheet.getDataRange().getValues();
  const candidates = data.slice(1).map(row => {
    try {
      // Extract file ID from the provided link format
      const url = row[0]; // Assuming the URL is in the first column
      const fileIdMatch = url.match(/\/d\/([^\/]*)\/view/);
      if (fileIdMatch) {
        const fileId = fileIdMatch[1];
        return {
          profilePic: `https://drive.google.com/uc?export=view&id=${fileId}`,
          name: row[1],
          email: row[2],
          resume: row[3],
          status: row[9]
        };
      } else {
        // Handle invalid URLs more gracefully (optional)
        console.warn(`Invalid file URL format in row: ${row[0]}`);
        return null; // or provide a default profile picture URL
      }
    } catch (error) {
      Logger.log(`Error processing row: ${error.message}`);
      return null; // or handle the error as needed
    }
  }).filter(candidate => candidate !== null);
  console.log(candidates)
  return candidates;
}

function getNumberOfApplicants() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('candidates');
  const data = sheet.getDataRange().getValues();
  return data.length - 1; // Subtract 1 to exclude the header row
}


//feedback
function copyToReport() {
  var sourceSpreadsheet = SpreadsheetApp.openById('1s8FpuG-An6Tl-aezOfRZCNM_4_dz_W6zcAqUCOPXaGc'); // ID of the spreadsheet with the form responses
  var sourceSheet = sourceSpreadsheet.getSheetByName('Interviewer feedback');
  if (!sourceSheet) {
    Logger.log("Source sheet 'Interviewer feedback' not found");
    return;
  }
  
  var sourceRange = sourceSheet.getDataRange();
  var sourceData = sourceRange.getValues();
  
  var destSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Destination Spreadsheet Name: " + destSpreadsheet.getName());
  var destSheet = destSpreadsheet.getSheetByName('InterviewerForm');
  
  if (!destSheet) {
    Logger.log("Destination sheet 'InterviewerForm' not found");
    return;
  }
  
  // Clear the destination sheet first (except for header row)
  destSheet.getRange(2, 1, destSheet.getMaxRows()-1, destSheet.getMaxColumns()).clearContent();
  
  // Copy the data from the source sheet to the destination sheet
  var numRows = sourceData.length - 1; // Number of rows excluding the header row
  var numCols = sourceData[0].length;  // Number of columns
  if (numRows > 0) {
    destSheet.getRange(2, 1, numRows, numCols).setValues(sourceData.slice(1));
  }
}


function mergeFeedbackData() {
    try {
        const spreadsheetId = '1-oiotcQz72_HqX9cqNDVzsHwLK0vKPFFJOMUcU3ksk8';
        
        // Access the sheets
        const ss = SpreadsheetApp.openById(spreadsheetId);
        const interviewSheet = ss.getSheetByName('Interview');
        const feedbackSheet = ss.getSheetByName('InterviewerForm');
        const mergeSheet = ss.getSheetByName('merge');
        
        // Check if sheets are properly retrieved
        if (!interviewSheet || !feedbackSheet || !mergeSheet) {
            throw new Error('One or more sheets could not be found. Please check the sheet names.');
        }
        
        // Fetch data from 'interview' sheet
        const interviewData = interviewSheet.getDataRange().getValues();
        
        // Fetch data from 'InterviewForm' sheet
        const feedbackData = feedbackSheet.getDataRange().getValues();
        
        // Prepare to store merged results
        const results = [];
        
        // Create header for merge sheet if needed
        if (mergeSheet.getLastRow() === 0) {
            mergeSheet.appendRow(['Candidate Name', 'Candidate Email', 'Interview Date', 'Interview Time', 'Remarks', 'General Comments', 'Satisfaction']);
        }
        
        // Map feedback data by Candidate Email
        const feedbackMap = new Map();
        for (let i = 1; i < feedbackData.length; i++) {
            const row = feedbackData[i];
            const email = row[4]; // Candidate Email is the 5th column (index 4)
            feedbackMap.set(email, {
                remarks: row[1], // Skills of candidates
                generalComments: row[2], // General comment of the interviewee
                satisfaction: row[3] // Rate your satisfaction of the interviewee(1-5)
            });
        }
        
        // Function to convert 12-hour time format to 24-hour time format
        function convertTo24HourFormat(timeStr) {
            if (timeStr instanceof Date) {
                // If timeStr is a Date object, format it to a string
                timeStr = Utilities.formatDate(timeStr, Session.getScriptTimeZone(), 'hh:mm:ss a');
            }
            const timeParts = timeStr.split(/[:\s]/);
            let hours = parseInt(timeParts[0], 10);
            const minutes = timeParts[1];
            const seconds = timeParts[2];
            const ampm = timeParts[3];
            
            if (ampm === 'PM' && hours < 12) hours += 12;
            if (ampm === 'AM' && hours === 12) hours = 0;
            
            return (hours < 10 ? '0' : '') + hours + ':' + minutes + ':' + seconds;
        }
        
        // Merge data
        for (let i = 1; i < interviewData.length; i++) {
            const row = interviewData[i];
            const email = row[1]; // Candidate Email is the 2nd column (index 1)
            
            if (feedbackMap.has(email)) {
                const feedbackRow = feedbackMap.get(email);
                const satisfactionRating = parseInt(feedbackRow.satisfaction);
                const stars = '★'.repeat(satisfactionRating) + '☆'.repeat(5 - satisfactionRating);
                
                // Convert Interview Time from 12-hour to 24-hour format
                const interviewTime = row[4];
                const formattedTime = convertTo24HourFormat(interviewTime);
                
                results.push([
                    row[0], // Candidate Name
                    email, // Candidate Email
                    row[3], // Interview Date
                    formattedTime, // Interview Time
                    feedbackRow.remarks, // Skills of candidates
                    feedbackRow.generalComments, // General comment of the interviewee
                    stars // Satisfaction as stars
                ]);
            }
        }
        
        // Clear existing data in merge sheet
        mergeSheet.clearContents();
        
        // Add headers and data
        mergeSheet.appendRow(['Candidate Name', 'Candidate Email', 'Interview Date', 'Interview Time', 'Remarks', 'General Comments', 'Satisfaction']);
        results.forEach(row => {
            mergeSheet.appendRow(row);
        });
        
        console.log('Data merged into sheet:', ss.getUrl()); // Use getUrl() on the Spreadsheet object
    } catch (error) {
        console.error('Error merging feedback data:', error.message);
    }
}




function getFeedbackDetails() {
    const sheetId = '1-oiotcQz72_HqX9cqNDVzsHwLK0vKPFFJOMUcU3ksk8';
    const sheetName = 'merge';

    const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
    if (!sheet) {
        Logger.log('Sheet not found:', sheetName);
        return { details: [] };
    }

    const data = sheet.getDataRange().getValues();
    Logger.log('Sheet Data: %s', JSON.stringify(data));  // Add this line for logging

    // Function to parse '1899-12-30TXX:XX:XX.000Z' date string to Date object
    function parseDateString(dateStr) {
        return new Date(dateStr);
    }

    // Skip the header row
    const details = data.slice(1).map(row => ({
        candidateName: row[0],
        candidateEmail: row[1],
        interviewDate: Utilities.formatDate(new Date(row[2]), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        interviewTime: Utilities.formatDate(parseDateString(row[3]), Session.getScriptTimeZone(), 'HH:mm:ss'),
        remarks: row[4],
        generalComments: row[5],
        satisfaction: row[6]
    }));
    
    Logger.log('Details: %s', JSON.stringify(details));  // Add this line for logging

    return { details };
}

function logFeedbackDetails() {
    const sheet = SpreadsheetApp.openById('1-oiotcQz72_HqX9cqNDVzsHwLK0vKPFFJOMUcU3ksk8').getSheetByName('merge');
    const data = sheet.getDataRange().getValues();
    
    // Skip header row and format data
    const details = data.slice(1).map(row => ({
        candidateName: row[0],
        candidateEmail: row[1],
        interviewDate: row[2],
        interviewTime: row[3],
        remarks: row[4],
        generalComments: row[5],
        satisfaction: row[6]
    }));
    
    // Log details to the Apps Script execution log
    details.forEach((detail, index) => {
        Logger.log(`Record ${index + 1}:`);
        Logger.log(`Candidate Name: ${detail.candidateName}`);
        Logger.log(`Candidate Email: ${detail.candidateEmail}`);
        Logger.log(`Interview Date: ${detail.interviewDate}`);
        Logger.log(`Interview Time: ${detail.interviewTime}`);
        Logger.log(`Remarks: ${detail.remarks}`);
        Logger.log(`General Comments: ${detail.generalComments}`);
        Logger.log(`Satisfaction: ${detail.satisfaction}`);
        Logger.log('---');
    });
}

function copyResponsesToReport() {
  var sourceSpreadsheet = SpreadsheetApp.openById('1QIhhaZcPJ4RRcAjzGhNqL5vYX3uvzIP8RNUFubI9j_0'); // ID of the spreadsheet with the form responses
  var sourceSheet = sourceSpreadsheet.getSheetByName('candidates feedback');
  if (!sourceSheet) {
    Logger.log("Source sheet 'candidates feedback' not found");
    return;
  }
  
  var sourceRange = sourceSheet.getDataRange();
  var sourceData = sourceRange.getValues();
  
  var destSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Destination Spreadsheet Name: " + destSpreadsheet.getName());
  var destSheet = destSpreadsheet.getSheetByName('FeedbackReport');
  
  if (!destSheet) {
    Logger.log("Destination sheet 'FeedbackReport' not found");
    return;
  }
  
  // Clear the destination sheet first (except for header row)
  destSheet.getRange(2, 1, destSheet.getMaxRows()-1, destSheet.getMaxColumns()).clearContent();
  
  // Copy the data from the source sheet to the destination sheet
  var numRows = sourceData.length - 1; // Number of rows excluding the header row
  var numCols = sourceData[0].length;  // Number of columns
  if (numRows > 0) {
    destSheet.getRange(2, 1, numRows, numCols).setValues(sourceData.slice(1));
  }
}

function categorizeResponses() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FeedbackReport');
    var data = sheet.getDataRange().getValues();
    var ratingColumnIndex = 1; // Assuming the rating is in the second column (index 1)
    var categoryColumnIndex = data[0].length; // New category column after the last existing column

    Logger.log("Categorization started");

    // Add header for category column if not present
    if (data[0].length <= categoryColumnIndex) {
      sheet.getRange(1, categoryColumnIndex + 1).setValue("Category");
    }

    for (var i = 1; i < data.length; i++) {
      if (data[i][ratingColumnIndex] != "") {
        var rating = parseInt(data[i][ratingColumnIndex], 10);
        if (rating >= 3) {
          sheet.getRange(i + 1, categoryColumnIndex + 1).setValue("Positive");
          Logger.log("Row " + (i + 1) + " categorized as Positive");
        } else {
          sheet.getRange(i + 1, categoryColumnIndex + 1).setValue("Negative");
          Logger.log("Row " + (i + 1) + " categorized as Negative");
        }
      }
    }
  } catch (error) {
    Logger.log("Error in categorizing responses: " + error.message);
  }
}

function getFeedbackData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FeedbackReport');
  var data = sheet.getDataRange().getValues();
  var positive = 0;
  var negative = 0;
  var categoryColumnIndex = data[0].length - 1; // Assuming the category column is the last column

  for (var i = 1; i < data.length; i++) {
    if (data[i][categoryColumnIndex] === "Positive") {
      positive++;
    } else if (data[i][categoryColumnIndex] === "Negative") {
      negative++;
    }
  }

  return { positive: positive, negative: negative };
}



function displayFeedbackData() {
  var feedbackData = getFeedbackData();
  Logger.log("Positive feedback: " + feedbackData.positive);
  Logger.log("Negative feedback: " + feedbackData.negative);
}

//display Profile Picture
function getProfilePic() {
  var fileId = '1v6Zxc8i-qAASQHQvTUN-4p6lVgOSCPxb'; // Replace with your actual file ID
  var file = DriveApp.getFileById(fileId);
  var blob = file.getBlob();
  var base64Data = Utilities.base64Encode(blob.getBytes());
  var contentType = blob.getContentType();
  
  // Return just the image data URL
  return "data:" + contentType + ";base64," + base64Data;
}

//schedule interview
function scheduleInterviewFromWeb(data) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.appendRow([data.candidateName, data.candidateEmail, data.interviewerEmail, data.interviewDate, data.interviewTime]);
  
    const lastRow = sheet.getLastRow();  // Get the index of the last row
    scheduleInterview(lastRow);  // Pass the index of the last row
}

function scheduleInterview(rowIndex) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const dataRange = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn());  // Get the range of the last row
    const data = dataRange.getValues()[0];  // Get the values of the last row
    const calendar = CalendarApp.getDefaultCalendar();

    const [candidateName, candidateEmail, interviewerEmail, interviewDate, interviewTime] = data;

    // Check if interviewDate is a Date object and convert it to a string in 'yyyy-mm-dd' format if needed
    let formattedDate = interviewDate;
    if (interviewDate instanceof Date) {
        const year = interviewDate.getFullYear();
        const month = String(interviewDate.getMonth() + 1).padStart(2, '0');
        const day = String(interviewDate.getDate()).padStart(2, '0');
        formattedDate = `${year}-${month}-${day}`;
    }

    // Check if interviewTime is a Date object and convert it to a string in 'HH:mm' format if needed
    let formattedTime = interviewTime;
    if (interviewTime instanceof Date) {
        const hours = String(interviewTime.getHours()).padStart(2, '0');
        const minutes = String(interviewTime.getMinutes()).padStart(2, '0');
        formattedTime = `${hours}:${minutes}`;
    }

    if (!candidateName || !candidateEmail || !interviewerEmail || !formattedDate || !formattedTime) {
        Logger.log(`Skipping row ${rowIndex}: Missing required data`);
        return;
    }

    try {
        const [year, month, day] = formattedDate.split('-');
        const [hour, minute] = formattedTime.split(':');
        const startDate = new Date(year, month - 1, day, hour, minute);
        const endDate = new Date(year, month - 1, day, hour, parseInt(minute) + 30);

        const event = calendar.createEvent(
            `Interview with ${candidateName}`,
            startDate,
            endDate,
            { guests: `${candidateEmail},${interviewerEmail}`, sendInvites: true }
        );

        sheet.getRange(rowIndex, 6).setValue(event.getId());

        const selectedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('selectedCandidates');
        const selectedData = selectedSheet.getDataRange().getValues();

        for (let i = 1; i < selectedData.length; i++) {
            if (selectedData[i][1] === candidateEmail) {
                // Assuming the columns are:
                // Name, Email, Status, Interviewer Email, Interview Date, Interview Time
                selectedSheet.getRange(i + 1, 3).setValue('Scheduled'); // Update status
                selectedSheet.getRange(i + 1, 4).setValue(interviewerEmail); // Update interviewer email
                selectedSheet.getRange(i + 1, 5).setValue(formattedDate); // Update interview date
                selectedSheet.getRange(i + 1, 6).setValue(formattedTime); // Update interview time
                break;
            }
        }
    } catch (error) {
        Logger.log(`Error processing row ${rowIndex}: ${error.message}`);
    }
}

function scheduleCandidateStatus(email) {
    updateCandidateScheduleStatus(email, 'Scheduled');
}

function updateCandidateScheduleStatus(email, status) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('selectedCandidates');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
        if (data[i][2] === email) {
            sheet.getRange(i + 1, 10).setValue(status);
            break;
        }
    }
}

//resume analysis
function resumeAnalysis() {
  convertPdfsToGoogleDocs();
}

function convertPdfsToGoogleDocs() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('selectedCandidates');
  var data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) { // Skip header row
    const resumeLink = data[i][3]; // Assuming the resume links are in the 4th column (index 3)
    const fileId = getFileIdFromUrl(resumeLink);
    const docId = convertPdfToGoogleDoc(fileId);
    
    var analysis = analyzeResume(docId);
    
    // Update the Google Sheet with the analysis results in separate columns
    sheet.getRange(i + 1, 5).setValue(analysis.workExperience.join(', '));
    sheet.getRange(i + 1, 6).setValue(analysis.education.bachelor + ', ' + analysis.education.university);
    sheet.getRange(i + 1, 7).setValue(analysis.expertise.join(', '));
    sheet.getRange(i + 1, 8).setValue(analysis.certifications.join(', '));
    sheet.getRange(i + 1, 9).setValue(analysis.numberOfAwards);
  }
}

function getFileIdFromUrl(url) {
  var fileIdMatch = url.match(/[-\w]{25,}/);
  return fileIdMatch ? fileIdMatch[0] : null;
}

function convertPdfToGoogleDoc(fileId) {
  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
  const newFileName = file.getName().replace(/\.pdf$/, '') + ' converted';
  const resource = {
    title: newFileName,
    mimeType: MimeType.GOOGLE_DOCS
  };
  const options = {
    ocr: true,
    ocrLanguage: 'en' // Specify OCR language; change as needed
  };
  const convertedFile = Drive.Files.create(resource, blob, options);
  const convertedFileObj = DriveApp.getFileById(convertedFile.id);
  convertedFileObj.setName(newFileName);
  return convertedFile.id;
}

function analyzeResume(docId) {
  var doc = DocumentApp.openById(docId);
  var body = doc.getBody();
  var text = body.getText();

  var educationPattern = /EDUCATION\s*([\s\S]*?)(?:EXPERTISE|$)/;
  var workExperiencePattern = /(?:PROFESSIONAL EXPERIENCE|WORK EXPERIENCE)\s*([\s\S]*?)(?:EDUCATION|$)/;
  var expertisePattern = /EXPERTISE\s*PROGRAMMING LANGUAGE USED\s*([\s\S]*?)(?:CERTIFICATION|$)/;
  var certificationPattern = /CERTIFICATION\s*([\s\S]*?)(?:AWARD|$)/;
  var awardPattern = /(?:AWARD|ACHIEVEMENT)\s*([\s\S]*?)(?:\n\n|$)/;

  var analysis = {
    workExperience: extractWorkExperience(text.match(workExperiencePattern)),
    education: extractEducation(text.match(educationPattern)),
    expertise: extractExpertise(text.match(expertisePattern)),
    certifications: extractCertifications(text.match(certificationPattern)),
    numberOfAwards: countAwards(text.match(awardPattern))
  };

  return analysis;
}

function extractEducation(match) {
  if (!match) return { bachelor: '', university: '' };

  var educationText = match[1].trim();
  var educationLines = educationText.split('\n').map(item => item.trim());

  var bachelor = educationLines[0];
  var university = educationLines[1];

  return { bachelor, university };
}

function extractWorkExperience(match) {
  if (!match) return [];

  var jobPattern = /([A-Za-z ]+),\s*([A-Za-z &]+)\s*\(([^)]+)\)/g;
  var workExperienceText = match[1].trim();
  var workExperienceList = [];

  var jobMatch;
  while ((jobMatch = jobPattern.exec(workExperienceText)) !== null) {
    var jobTitle = jobMatch[1].trim();
    var company = jobMatch[2].trim();
    var dateRange = jobMatch[3].trim();
    workExperienceList.push(`${jobTitle}, ${company} (${dateRange})`);
  }

  return workExperienceList;
}

function extractExpertise(match) {
  if (!match) return [];

  var expertiseText = match[1].trim();
  var expertiseList = expertiseText.split('\n').map(item => item.trim());

  return expertiseList;
}

function extractCertifications(match) {
  if (!match) return [];

  var certificationText = match[1].trim();
  var certificationList = certificationText.split('\n').map(item => item.trim());

  return certificationList;
}

function countAwards(match) {
  if (!match) return 0;

  var awardText = match[1].trim();
  var awardList = awardText.split('\n').map(item => item.trim());

  return awardList.length;
}

//RESUME chart
// Function to fetch Pie Chart data from Google Sheets
function getSelectedPieChartData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Charts");
  if (!sheet) {
    Logger.log("Sheet 'Charts' not found.");
    return;
  }
  
  var range = sheet.getRange("I1:J" + sheet.getLastRow());
  var values = range.getValues();
  
  return values;
}

// Function to fetch Bar Chart data for Number of Awards
function getAwardsBarChartData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Charts");
  if (!sheet) {
    Logger.log("Sheet 'Charts' not found.");
    return;
  }
  
  var range = sheet.getRange("F1:G" + sheet.getLastRow());
  var values = range.getValues();
  
  return values;
}

// Function to fetch Column Chart data for Certifications
function getCertificationsColumnChartData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Charts");
  if (!sheet) {
    Logger.log("Sheet 'Charts' not found.");
    return;
  }
  
  var range = sheet.getRange("L1:M" + sheet.getLastRow());
  var values = range.getValues();
  
  return values;
}

function generateCharts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("selectedCandidates");
  if (!sheet) {
    Logger.log("Sheet 'selectedCandidate' not found.");
    return;
  }
  
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  
  // Try to delete the existing 'Charts' sheet if it exists, to start fresh
  var existingChartSheet = ss.getSheetByName('Charts');
  if (existingChartSheet) ss.deleteSheet(existingChartSheet);
  
  // Create a new sheet for charts
  var chartSheet = ss.insertSheet('Charts');
  

  // 2. Bar Chart for Number of Awards
  var awardsData = [['Name', 'Number of Awards']];
  data.slice(1).forEach(row => {
    awardsData.push([row[1], parseInt(row[8]) || 0]); // Ensure conversion to integer and handle non-integer values gracefully
  });
  var awardsRange = chartSheet.getRange(1, 6, awardsData.length, 2); // Adjusted for clarity
  awardsRange.setValues(awardsData);

  var awardsChart = chartSheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(awardsRange)
      .setPosition(10, 6, 0, 0) // Adjusted position for clarity
      .build();
  chartSheet.insertChart(awardsChart);

  // 3. Pie Chart for Expertise Distribution
  var expertise = {};
  data.slice(1).forEach(row => {
    var exps = row[6].split(', ');
    exps.forEach(exp => {
      expertise[exp] = (expertise[exp] || 0) + 1;
    });
  });
  
  var expertiseData = [['Expertise', 'Count']];
  for (var exp in expertise) {
    expertiseData.push([exp, expertise[exp]]);
  }
  
  var expertiseRange = chartSheet.getRange(1, 9, expertiseData.length, 2); // Adjusted for clarity
  expertiseRange.setValues(expertiseData);

  var expertiseChart = chartSheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(expertiseRange)
      .setPosition(20, 9, 0, 0) // Adjusted position for clarity
      .build();
  chartSheet.insertChart(expertiseChart);

  // 4. Column Chart for Certifications
  var certifications = {};
  data.slice(1).forEach(row => {
    var certs = row[7].split(', ');
    certs.forEach(cert => {
      certifications[cert] = (certifications[cert] || 0) + 1;
    });
  });

  var certificationsData = [['Certification', 'Count']];
  for (var cert in certifications) {
    certificationsData.push([cert, certifications[cert]]);
  }

  var certificationsRange = chartSheet.getRange(1, 12, certificationsData.length, 2); // Adjusted for clarity
  certificationsRange.setValues(certificationsData);

  var certificationsChart = chartSheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(certificationsRange)
      .setPosition(30, 12, 0, 0) // Adjusted position for clarity
      .build();
  
  chartSheet.insertChart(certificationsChart);

  // 5. Table for Education Background
  var educationData = [['Name', 'Education']];
data.slice(1).forEach(row => {
  educationData.push([row[1], row[5]]);
});
var educationRange = chartSheet.getRange(1, 15, educationData.length, 2); // Adjusted for clarity
educationRange.setValues(educationData);

// Optionally, format the range to look more like a table
educationRange.setBorder(true, true, true, true, true, true);
educationRange.setBackground('#f0f0f0');
chartSheet.setColumnWidths(15, 2, 150); // Adjust column widths for readability
}
