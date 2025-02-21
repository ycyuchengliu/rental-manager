function onOpen(e) {
     SpreadsheetApp.getUi().createMenu('Refresh').addItem('Refresh', 'distributeUtilityCostsAndSendEmails').addToUi();
   }

function distributeUtilityCostsAndSendEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var utilitiesSheet = ss.getSheetByName("Utility Costs");
  var rentersSheet = ss.getSheetByName("Tenants");
  var resultsSheet = ss.getSheetByName("Expenses Breakdown") || ss.insertSheet("Expenses Breakdown");

  // Step 1: Read Utilities Data
  var utilitiesData = utilitiesSheet.getDataRange().getValues();
  utilitiesData.shift(); // Remove headers

  // Step 2: Read Renters Data (Now includes emails)
  var rentersData = rentersSheet.getDataRange().getValues();
  var renterEmails = {};
  rentersData.shift(); // Remove headers
  rentersData.forEach(function(row) {
    var [renterName, leaseStart, leaseEnd, email] = row;
    renterEmails[renterName] = email; // Store renter email
  });

  // Step 3: Prepare Results Sheet
  resultsSheet.clear();
  resultsSheet.appendRow(["Tenant Name", "Billing Period", "Days in Period", "Type", "Utility Cost"]);

  var results = {};

  // Step 4: Process Each Utility Bill
  utilitiesData.forEach(function(utility) {
    var [billID, utilityType, amount, startDate, endDate] = utility;
    startDate = new Date(startDate);
    endDate = new Date(endDate);

    // Format Billing Period as "YYYY-MM-DD to YYYY-MM-DD (Utility Type)"
    var billingPeriod = formatDate(startDate) + " to " + formatDate(endDate) + " (" + utilityType + ")";

    var totalDaysInBillingPeriod = Math.round((endDate - startDate) / (1000 * 60 * 60 * 24)) + 1;

    // Step 5: Find Renters Active During the Billing Period
    var activeRenters = rentersData.map(function(renter) {
      var [renterName, leaseStart, leaseEnd, email] = renter;
      leaseStart = new Date(leaseStart);
      leaseEnd = new Date(leaseEnd);

      // Calculate overlapping days
      var overlapStart = new Date(Math.max(leaseStart, startDate));
      var overlapEnd = new Date(Math.min(leaseEnd, endDate));
      var daysInPeriod = Math.round((overlapEnd - overlapStart) / (1000 * 60 * 60 * 24)) + 1;

      return (daysInPeriod > 0) ? { renterName, daysInPeriod, email } : null;
    }).filter(renter => renter !== null); // Remove renters with zero days

    var totalDaysOccupied = activeRenters.reduce((sum, renter) => sum + renter.daysInPeriod, 0);

    // Step 6: Allocate Cost Proportionally to Days Stayed
    if (totalDaysOccupied > 0) {
      activeRenters.forEach(function(renter) {
        var { renterName, daysInPeriod, email } = renter;
        var renterKey = renterName + "-" + billingPeriod + "-" + utilityType;

        if (!results[renterKey]) {
          results[renterKey] = { renterName, billingPeriod, daysInPeriod, utilityType, UtilityCost: 0};
        }

        var renterShare = (daysInPeriod / totalDaysOccupied) * amount;
        results[renterKey]["UtilityCost"] += renterShare;
      });
    }
  });

  // Step 7: Write Results to the Sheet & Send Emails
  Object.values(results).forEach(row => {
    resultsSheet.appendRow([
      row.renterName,
      row.billingPeriod,
      Math.round(row.daysInPeriod), // Ensure Days in Period is an integer
      row.utilityType,
      formatCurrency(row.UtilityCost)
    ]);
  });

  Logger.log("Utility costs distributed and emails sent!");
}

// Helper function to format date as YYYY-MM-DD
function formatDate(date) {
  return date.getFullYear() + "-" + String(date.getMonth() + 1).padStart(2, "0") + "-" + String(date.getDate()).padStart(2, "0");
}

// Helper function to format currency as $X.XX
function formatCurrency(amount) {
  return "$" + amount.toFixed(2);
}


//not in use
// Helper function to send email notifications
function sendEmailNotification(renterName, email, billingPeriod, days, utilityType, utilityCost, total) {
  var subject = "Your Utility Cost Breakdown";
  var body = "Hello " + renterName + ",\n\n" +
             "Here is your utility cost breakdown for " + billingPeriod + ":\n\n" +
             "Days in Billing Period: " + days + "\n" +
             "Utility Type: " + utilityType + "\n" +
             "Utility Cost: " + formatCurrency(utilityCost) + "\n" +
             "Total Cost: " + formatCurrency(total) + "\n\n" +
             "Please reach out if you have any questions.\n\n" +
             "Best,\n[Your Name or Management Team]";

  MailApp.sendEmail(email, subject, body);
}
