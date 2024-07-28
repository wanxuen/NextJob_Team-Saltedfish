function scheduleInterviewFromWeb(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.appendRow([data.candidateName, data.candidateEmail, data.interviewerEmail, data.interviewDate, data.interviewTime]);
  scheduleInterview();
}

function scheduleInterview() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const calendar = CalendarApp.getDefaultCalendar();

  data.forEach((row, index) => {
    if (index === 0) return; // Skip header row

    const [candidateName, candidateEmail, interviewerEmail, interviewDate, interviewTimeString] = row;
    Logger.log(`interviewTimeString data type: ${typeof interviewTimeString}`);

          // Check if interviewTimeString is a date object
      if (interviewTimeString instanceof Date) {
        // Extract time using getHours() and getMinutes()
        const hour = interviewTimeString.getHours();
        const minute = interviewTimeString.getMinutes();
        interviewTime = `${hour}:${minute}`; // Formatted time string
      }

    if (!candidateName || !candidateEmail || !interviewerEmail || !interviewDate || !interviewTime) {
      Logger.log(`Skipping row ${index + 1}: Missing required data`);
      return;
    }

    try {
      const [day, month, year] = interviewDate.split('/');
      const [hour, minute] = interviewTime.split(':');
      const startDate = new Date(year, month - 1, day, hour, minute);
      const endDate = new Date(year, month - 1, day, hour, minute + 30);

      const event = calendar.createEvent(
        `Interview with ${candidateName}`,
        startDate,
        endDate,
        { guests: `${candidateEmail},${interviewerEmail}`, sendInvites: true }
      );

      sheet.getRange(index + 1, 6).setValue(event.getId());
    } catch (error) {
      Logger.log(`Error processing row ${index + 1}: ${error.message}`);
    }
  });
}

function generateRecruitmentReportFromWeb() {
  const candidateFeedbackSheet = SpreadsheetApp.openById('1QIhhaZcPJ4RRcAjzGhNqL5vYX3uvzIP8RNUFubI9j_0').getSheetByName('candidates feedback');
  const interviewerFeedbackSheet = SpreadsheetApp.openById('1s8FpuG-An6Tl-aezOfRZCNM_4_dz_W6zcAqUCOPXaGc').getSheetByName('interviewer feedback');
  const candidateFeedback = candidateFeedbackSheet.getDataRange().getValues();
  const interviewerFeedback = interviewerFeedbackSheet.getDataRange().getValues();

  let reportHtml = '<table border="1"><tr><th>Metric</th><th>Value</th></tr>';

  const totalInterviews = candidateFeedback.length - 1;
  const avgCandidateSatisfaction = candidateFeedback.reduce((sum, row, index) => index === 0 ? sum : sum + Number(row[1]), 0) / totalInterviews;
  const avgInterviewerSatisfaction = interviewerFeedback.reduce((sum, row, index) => index === 0 ? sum : sum + Number(row[1]), 0) / totalInterviews;

  reportHtml += `<tr><td>Total Interviews</td><td>${totalInterviews}</td></tr>`;
  reportHtml += `<tr><td>Average Candidate Satisfaction</td><td>${avgCandidateSatisfaction}</td></tr>`;
  reportHtml += `<tr><td>Average Interviewer Satisfaction</td><td>${avgInterviewerSatisfaction}</td></tr>`;
  reportHtml += '</table>';

  return reportHtml;
}
