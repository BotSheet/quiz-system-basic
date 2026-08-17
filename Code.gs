/**
 * Chemistry Weekly Quiz System - Version 1.0
 * Backend Server Script
 */

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('Chemistry Weekly Quiz')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getDb() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSettingsMap() {
  var sheet = getDb().getSheetByName('Settings');
  var data = sheet.getDataRange().getValues();
  var settings = {};
  for (var i = 1; i < data.length; i++) {
    var key = String(data[i][0]).trim();
    var val = String(data[i][1]).trim();
    if (key) {
      settings[key] = val;
    }
  }
  return settings;
}

function getSheetDataAsObjects(sheetName) {
  var sheet = getDb().getSheetByName(sheetName);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var result = [];
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = row[j];
    }
    result.push(obj);
  }
  return result;
}

function validateAndFetchTest(studentIdInput, emailInput) {
  try {
    var studentId = String(studentIdInput).trim();
    var email = String(emailInput).trim().toLowerCase();
    
    if (!studentId || !email) {
      return { success: false, message: 'Please provide both Student ID and registered Email.' };
    }

    var settings = getSettingsMap();
    var activeTestId = settings['Active Test ID'];
    var teacherName = settings['Teacher Name'] || 'Chemistry Department';

    if (!activeTestId) {
      return { success: false, message: 'No active test is currently scheduled by the teacher.' };
    }

    // 1. Verify Student
    var students = getSheetDataAsObjects('Students');
    var matchedStudent = null;
    for (var i = 0; i < students.length; i++) {
      if (String(students[i]['Student ID']).trim().toUpperCase() === studentId.toUpperCase()) {
        matchedStudent = students[i];
        break;
      }
    }

    if (!matchedStudent) {
      return { success: false, message: 'Student ID not found in records.' };
    }

    if (String(matchedStudent['Email']).trim().toLowerCase() !== email) {
      return { success: false, message: 'Email address does not match our registered records.' };
    }

    if (matchedStudent['Active'] !== true && String(matchedStudent['Active']).toUpperCase() !== 'TRUE') {
      return { success: false, message: 'Your student account is currently inactive. Contact your teacher.' };
    }

    // 2. Verify Test
    var tests = getSheetDataAsObjects('Tests');
    var matchedTest = null;
    for (var j = 0; j < tests.length; j++) {
      if (String(tests[j]['Test ID']).trim() === activeTestId) {
        matchedTest = tests[j];
        break;
      }
    }

    if (!matchedTest) {
      return { success: false, message: 'Active Test configuration not found.' };
    }

    if (String(matchedTest['Status']).trim().toLowerCase() !== 'published') {
      return { success: false, message: 'The current test is not open for submissions.' };
    }

    // 3. Verify Duplicate Submission
    var results = getSheetDataAsObjects('Results');
    for (var k = 0; k < results.length; k++) {
      var rTestId = String(results[k]['Test ID']).trim();
      var rStudentId = String(results[k]['Student ID']).trim().toUpperCase();
      if (rTestId === activeTestId && rStudentId === studentId.toUpperCase()) {
        return { success: false, message: 'You have already completed and submitted this test.' };
      }
    }

    // 4. Fetch Questions Mapped to this Test
    var testQuestions = getSheetDataAsObjects('TestQuestions');
    var targetQuestionIds = [];
    for (var l = 0; l < testQuestions.length; l++) {
      if (String(testQuestions[l]['Test ID']).trim() === activeTestId) {
        targetQuestionIds.push(String(testQuestions[l]['Question ID']).trim());
      }
    }

    if (targetQuestionIds.length === 0) {
      return { success: false, message: 'No questions have been assigned to this test.' };
    }

    var allQuestions = getSheetDataAsObjects('Questions');
    var sanitizedQuestions = [];

    for (var m = 0; m < targetQuestionIds.length; m++) {
      var qId = targetQuestionIds[m];
      for (var n = 0; n < allQuestions.length; n++) {
        if (String(allQuestions[n]['Question ID']).trim() === qId) {
          var q = allQuestions[n];
          // Strip correct answers, explanations, and internal fields before sending to client
          sanitizedQuestions.push({
            questionId: String(q['Question ID']).trim(),
            chapter: q['Chapter'],
            topic: q['Topic'],
            question: q['Question'],
            optionA: q['Option A'],
            optionB: q['Option B'],
            optionC: q['Option C'],
            optionD: q['Option D'],
            marks: Number(q['Marks']) || 1,
            negativeMarks: Number(q['Negative Marks']) || 0
          });
          break;
        }
      }
    }

    return {
      success: true,
      data: {
        testId: activeTestId,
        testName: matchedTest['Test Name'],
        durationMinutes: Number(matchedTest['Duration']) || 20,
        student: {
          studentId: matchedStudent['Student ID'],
          name: matchedStudent['Name'],
          email: matchedStudent['Email'],
          studentClass: matchedStudent['Class'],
          batch: matchedStudent['Batch']
        },
        questions: sanitizedQuestions
      }
    };

  } catch (err) {
    Logger.log('Error in validateAndFetchTest: ' + err.toString());
    return { success: false, message: 'A server error occurred while preparing your test. Please try again.' };
  }
}

function submitQuizAnswers(payload) {
  try {
    var studentId = String(payload.studentId).trim();
    var testId = String(payload.testId).trim();
    var answers = payload.answers || {}; // Format: { "Q001": "A", "Q002": "C" }
    var timeTakenSeconds = payload.timeTakenSeconds || 0;

    var ss = getDb();
    var resultsSheet = ss.getSheetByName('Results');
    var detailsSheet = ss.getSheetByName('AnswerDetails');

    // Duplicate Check on Submit
    var existingResults = resultsSheet.getDataRange().getValues();
    for (var i = 1; i < existingResults.length; i++) {
      if (String(existingResults[i][1]).trim() === testId && 
          String(existingResults[i][2]).trim().toUpperCase() === studentId.toUpperCase()) {
        return { success: false, message: 'Submission rejected: Duplicate attempt detected.' };
      }
    }

    // Load necessary sheets for grading
    var students = getSheetDataAsObjects('Students');
    var matchedStudent = null;
    for (var s = 0; s < students.length; s++) {
      if (String(students[s]['Student ID']).trim().toUpperCase() === studentId.toUpperCase()) {
        matchedStudent = students[s];
        break;
      }
    }

    var tests = getSheetDataAsObjects('Tests');
    var matchedTest = null;
    for (var t = 0; t < tests.length; t++) {
      if (String(tests[t]['Test ID']).trim() === testId) {
        matchedTest = tests[t];
        break;
      }
    }

    var testQuestions = getSheetDataAsObjects('TestQuestions');
    var targetQuestionIds = [];
    for (var tq = 0; tq < testQuestions.length; tq++) {
      if (String(testQuestions[tq]['Test ID']).trim() === testId) {
        targetQuestionIds.push(String(testQuestions[tq]['Question ID']).trim());
      }
    }

    var allQuestions = getSheetDataAsObjects('Questions');
    var qMap = {};
    for (var q = 0; q < allQuestions.length; q++) {
      var item = allQuestions[q];
      qMap[String(item['Question ID']).trim()] = item;
    }

    var totalScore = 0;
    var maxMarks = 0;
    var correctCount = 0;
    var incorrectCount = 0;
    var unansweredCount = 0;

    var topicStats = {}; // { topicName: { total: 0, correct: 0 } }
    var detailedRowsToAppend = [];
    var timestamp = new Date();

    for (var j = 0; j < targetQuestionIds.length; j++) {
      var qId = targetQuestionIds[j];
      var questionRecord = qMap[qId];
      if (!questionRecord) continue;

      var qMarks = Number(questionRecord['Marks']) || 1;
      var qNegMarks = Number(questionRecord['Negative Marks']) || 0;
      var correctAns = String(questionRecord['Correct Answer']).trim().toUpperCase();
      var topicName = String(questionRecord['Topic'] || questionRecord['Chapter'] || 'General Chemistry').trim();

      maxMarks += qMarks;

      if (!topicStats[topicName]) {
        topicStats[topicName] = { total: 0, correct: 0 };
      }
      topicStats[topicName].total += 1;

      var selected = answers[qId] ? String(answers[qId]).trim().toUpperCase() : '';
      var resultStatus = '';

      if (!selected) {
        unansweredCount++;
        resultStatus = 'Unanswered';
      } else if (selected === correctAns) {
        correctCount++;
        totalScore += qMarks;
        topicStats[topicName].correct += 1;
        resultStatus = 'Correct';
      } else {
        incorrectCount++;
        totalScore -= qNegMarks;
        resultStatus = 'Incorrect';
      }

      detailedRowsToAppend.push([
        timestamp,
        testId,
        studentId,
        qId,
        selected || 'None',
        correctAns,
        resultStatus
      ]);
    }

    var attemptedCount = correctCount + incorrectCount;
    var percentage = maxMarks > 0 ? ((totalScore / maxMarks) * 100) : 0;
    percentage = Math.max(0, Number(percentage.toFixed(2))); // Prevent negative percentages on report
    var accuracy = attemptedCount > 0 ? Number(((correctCount / attemptedCount) * 100).toFixed(2)) : 0;

    // Time Taken Formatting
    var minutes = Math.floor(timeTakenSeconds / 60);
    var seconds = timeTakenSeconds % 60;
    var timeTakenFormatted = minutes + ' min ' + (seconds < 10 ? '0' : '') + seconds + ' sec';

    // Topic Performance Breakdown
    var topicSummary = [];
    var highestAcc = -1;
    var lowestAcc = 101;
    var strongArea = 'N/A';
    var needsRevision = 'N/A';

    for (var top in topicStats) {
      var stats = topicStats[top];
      var topAcc = stats.total > 0 ? Number(((stats.correct / stats.total) * 100).toFixed(1)) : 0;
      topicSummary.push({
        topic: top,
        questions: stats.total,
        correct: stats.correct,
        accuracy: topAcc
      });

      if (stats.total >= 1) {
        if (topAcc > highestAcc) {
          highestAcc = topAcc;
          strongArea = top + ' (' + topAcc + '%)';
        }
        if (topAcc < lowestAcc) {
          lowestAcc = topAcc;
          needsRevision = top + ' (' + topAcc + '%)';
        }
      }
    }

    // Performance Feedback Message
    var feedbackMessage = '';
    if (percentage >= 90) {
      feedbackMessage = 'Excellent performance. Keep maintaining this level.';
    } else if (percentage >= 75) {
      feedbackMessage = 'Good performance. Revise the questions you missed.';
    } else if (percentage >= 50) {
      feedbackMessage = 'Good attempt. Focus on the topics where you made mistakes.';
    } else {
      feedbackMessage = 'Use this test to identify your weak areas and revise the relevant topics.';
    }

    // Save to Google Sheets
    var settings = getSettingsMap();
    var teacherName = settings['Teacher Name'] || 'Chemistry Faculty';
    var reportSent = 'No';

    // Append AnswerDetails
    if (detailedRowsToAppend.length > 0) {
      var detailsRange = detailsSheet.getRange(
        detailsSheet.getLastRow() + 1, 
        1, 
        detailedRowsToAppend.length, 
        detailedRowsToAppend[0].length
      );
      detailsRange.setValues(detailedRowsToAppend);
    }

    // Send Email
    var studentEmail = matchedStudent ? matchedStudent['Email'] : '';
    var studentName = matchedStudent ? matchedStudent['Name'] : 'Student';
    var testName = matchedTest ? matchedTest['Test Name'] : testId;
    var studentClass = matchedStudent ? matchedStudent['Class'] : '';
    var studentBatch = matchedStudent ? matchedStudent['Batch'] : '';

    if (studentEmail) {
      try {
        var emailSubject = 'Chemistry Weekly Test Report — ' + testName;
        var emailBody = 
          'Dear ' + studentName + ',\n\n' +
          'Your Chemistry Weekly Test has been evaluated.\n\n' +
          'Test: ' + testName + '\n' +
          'Class: ' + studentClass + '\n' +
          'Score: ' + totalScore + ' / ' + maxMarks + '\n' +
          'Percentage: ' + percentage + '%\n' +
          'Correct: ' + correctCount + '\n' +
          'Incorrect: ' + incorrectCount + '\n' +
          'Unanswered: ' + unansweredCount + '\n' +
          'Accuracy: ' + accuracy + '%\n' +
          'Time Taken: ' + timeTakenFormatted + '\n\n' +
          'Strong Area:\n' + strongArea + '\n\n' +
          'Needs Revision:\n' + needsRevision + '\n\n' +
          feedbackMessage + '\n\n' +
          'Keep practising and improving.\n\n' +
          'Regards,\n' +
          teacherName;

        MailApp.sendEmail(studentEmail, emailSubject, emailBody);
        reportSent = 'Yes';
      } catch (mailErr) {
        Logger.log('Email delivery failed: ' + mailErr.toString());
        reportSent = 'Failed: ' + mailErr.message;
      }
    }

    // Append Results
    resultsSheet.appendRow([
      timestamp,
      testId,
      studentId,
      studentName,
      studentEmail,
      studentClass,
      studentBatch,
      totalScore,
      maxMarks,
      percentage,
      correctCount,
      incorrectCount,
      unansweredCount,
      accuracy,
      timeTakenFormatted,
      reportSent
    ]);

    return {
      success: true,
      report: {
        studentName: studentName,
        testName: testName,
        score: totalScore,
        maxMarks: maxMarks,
        percentage: percentage,
        correct: correctCount,
        incorrect: incorrectCount,
        unanswered: unansweredCount,
        accuracy: accuracy,
        timeTaken: timeTakenFormatted,
        feedbackMessage: feedbackMessage,
        strongArea: strongArea,
        needsRevision: needsRevision,
        topicSummary: topicSummary
      }
    };

  } catch (err) {
    Logger.log('Error during test submission: ' + err.toString());
    return { success: false, message: 'Server error while calculating score: ' + err.toString() };
  }
}

function getTeacherDashboardData() {
  try {
    var settings = getSettingsMap();
    var activeTestId = settings['Active Test ID'];
    var results = getSheetDataAsObjects('Results');
    var filteredResults = [];

    for (var i = 0; i < results.length; i++) {
      if (String(results[i]['Test ID']).trim() === activeTestId) {
        filteredResults.push({
          studentName: results[i]['Student Name'],
          score: Number(results[i]['Score']) || 0,
          maxMarks: Number(results[i]['Maximum Marks']) || 0,
          percentage: Number(results[i]['Percentage']) || 0,
          correct: Number(results[i]['Correct']) || 0,
          incorrect: Number(results[i]['Incorrect']) || 0,
          accuracy: Number(results[i]['Accuracy']) || 0,
          timeTaken: results[i]['Time Taken']
        });
      }
    }

    var totalStudents = filteredResults.length;
    var avgScore = 0;
    var highestScore = 0;
    var lowestScore = 0;

    if (totalStudents > 0) {
      var scoreSum = 0;
      highestScore = filteredResults[0].score;
      lowestScore = filteredResults[0].score;

      for (var j = 0; j < filteredResults.length; j++) {
        var s = filteredResults[j].score;
        scoreSum += s;
        if (s > highestScore) highestScore = s;
        if (s < lowestScore) lowestScore = s;
      }
      avgScore = Number((scoreSum / totalStudents).toFixed(2));
    }

    return {
      success: true,
      activeTestId: activeTestId,
      totalCompleted: totalStudents,
      avgScore: avgScore,
      highestScore: highestScore,
      lowestScore: lowestScore,
      results: filteredResults
    };
  } catch (err) {
    Logger.log('Error fetching teacher stats: ' + err.toString());
    return { success: false, message: err.toString() };
  }
}
