/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import OpenAI from 'openai';

/* global document, Office */
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("app-body").onload = load_main_data;
    document.getElementById("btn-summarize").onclick = run;
    document.getElementById("btn-sync-data").onclick = syncData; 

    const item = Office.context.mailbox.item;
    if (item.itemType === Office.MailboxEnums.ItemType.Message) {
      document.getElementById("item-subject").innerText = item.subject;
      document.getElementById("item-subject-date").innerText = item.dateTimeCreated.format("dd-MM-yyyy");
    }
  }
});


async function load_main_data(){
  const item = Office.context.mailbox.item;
    if (item.itemType === Office.MailboxEnums.ItemType.Message) {
      document.getElementById("item-subject").innerText = item.subject;
    }
}

async function run() {
  const item = Office.context.mailbox.item;
  let emailData = {
    subject: item.subject || "No Subject",
    from: item.from && item.from.emailAddress ? item.from.emailAddress : "No Sender",
    to: "",
    cc: "",
    body: "",
    actions: []
  };

  if (item.itemType === Office.MailboxEnums.ItemType.Message) {
    // Assuming toRecipients and ccRecipients are directly accessible (synchronously)
    var emailDataArr = (item.to || []).map(recipient => recipient.emailAddress)

    emailData.to = emailDataArr.join(", ");

    var ccRecipientsArr = (item.cc || []).map(recipient => recipient.emailAddress)
    emailData.cc = ccRecipientsArr.join(", ");

    // Use getAsync method to get the body asynchronously
    await new Promise((resolve, reject) => {
      item.body.getAsync(Office.CoercionType.Text, function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          emailData.body = result.value;
          resolve();
        } else {
          console.error("Error getting item body:", result.error);
          emailData.body = "Error retrieving body";
          reject();
        }
      });
    });
  }
  console.log(emailData);
    // Continue with summarizing and displaying after fetching body
    const summary = await summarizeContent(emailData.subject + " " + emailData.body); // Simulated function
    emailData.actions = summary.actions;
  let recipients = emailDataArr.concat(ccRecipientsArr)  
  populateActions(emailData.actions, recipients);

  // displayEmailData(emailData); // Function to update the UI with the fetched data
}

function displayEmailData(emailData) {
  // Update the UI with email data and summarized actions
  // This includes filling in subject, from, to, cc, and the body as before
  // Additionally, render the actions in an editable format (e.g., a contenteditable div or a list of input fields)

  // For example, display actions:
  let actionListHTML = emailData.actions.map((action, index) => `<label>Action ${index+1}:</label>&nbsp;<input type="text" id="action-${index}" value="${action}">`).join("<br/>");
  document.getElementById("action-list").innerHTML = actionListHTML;
}

function gatherEmailData() {
  const item = Office.context.mailbox.item;
  let emailData = {
    subject: item.subject,
    from: item.from && item.from.emailAddress,
    toRecipients: [],
    ccRecipients: [],
    body: ""
  };

  // Assuming item is fully loaded; otherwise, you may need to load properties explicitly
  if (item.toRecipients) {
    emailData.toRecipients = item.toRecipients.map(recipient => recipient.emailAddress).join(", ");
  }

  if (item.ccRecipients) {
    emailData.ccRecipients = item.ccRecipients.map(recipient => recipient.emailAddress).join(", ");
  }

  // Accessing the body requires an asynchronous call
  item.body.getAsync(Office.CoercionType.Text, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      emailData.body = result.value;
      // Here you would continue processing now that you have the body
      console.log(emailData); // Example of next step
    }
  });

  // Since body fetching is async, this might not include body immediately
  return emailData;
}


async function syncData() {
  // Gathering basic email data
  // Assuming gatherEmailData is an asynchronous function
  console.log('Fetching email data...');
  const emailData = gatherEmailData(); // Use await to wait for the promise to resolve

  console.log(`Here is the gathered email data:`, emailData);

  // Gathering updated actions and their assignments
  const actionsTable = document.getElementById('actionsTable').getElementsByTagName('tbody')[0];
  const rows = actionsTable.rows;
  const tasks = Array.from(rows).map((row, index) => {
    const action = row.cells[0].getElementsByTagName("input")[0].value;
    const assignedTo = row.cells[1].getElementsByTagName("select")[0].value;
    return { Title: action, AssignedTo: assignedTo };
  });

  // Prepare attachments (modify this part according to how you manage attachments in your UI)
  const attachments = []; // This should be an array of objects with { AttachmentName, Content }

  const syncedData = {
    subject: emailData.subject,
    from: emailData.from,
    to: emailData.to,
    cc: emailData.cc,
    body: emailData.body,
    tasks: tasks,
    attachments: attachments
  };

  console.log(syncedData);
  // Call the function to sync data, passing the structured email data
  try {
    const result = await syncEmailData(syncedData); // Assuming this function is properly defined elsewhere
    alert("Data synced successfully!");
    // Handle success response
  } catch (error) {
    alert("Failed to sync data. Please try again later.");
    console.error("Failed to sync data", error);
    // Handle errors
  }
}

function showModal(message) {
  document.querySelector('#syncResultModal .modal-body').textContent = message;
  $('#syncResultModal').modal('show'); // Bootstrap way to show modal
}




// Mock function to simulate summarizing content
async function summarizeContent(content) {
  try {
    const openai = new OpenAI({
      apiKey: '',
      dangerouslyAllowBrowser: true
    });
    // Hypothetical correct method for SDK version you're using
    const response = await openai.chat.completions.create({
      model: "gpt-3.5-turbo", // or "gpt-4" as available
      // prompt: `Summarize this into actionable tasks:\n\n${content}`,
      temperature: 0,
      max_tokens: 1000,
      messages: [{ role: "system", content: `Summarize the following content into actionable tasks and don't get out of the given context. Don't show "Tasks" as a title at the begining while summarizaing the actions. Return only the list of actions directly :\n\n${content}` }]
    });

    // Process the response to extract actions
    // This part remains similar, assuming the response structure is consistent
    const summaryText = response.choices[0].message.content.trim();
    const actions = summaryText.split("\n").filter(line => line.length > 0);

    return { actions };
  } catch (error) {
    console.error("Error summarizing content:", error);
    return { actions: ["Error summarizing content"] };
  }
}

// Mock function to simulate syncing data to a web service
async function mockSyncDataService(data) {
  // Simulate an API call to sync data
  return { success: true }; // Mocked sync success
}

async function syncEmailData(emailData) {
  const { subject, from, to, cc, body, tasks, attachments } = emailData;

  // Prepare the subject object
  const subjectObject = {
      Title: subject,
      CreatedOn: new Date().toISOString(), // Assuming current date as creation date
      From: from,
      To: to,
      CC: cc,
      Body: body,
      Tasks: tasks,
      Attachments: attachments
  };

  // Make the HTTP request to the API
  try {
      const response = await fetch('https://4tv76xgc-5215.inc1.devtunnels.ms/api/subjects', {
          method: 'POST',
          headers: {
              'Content-Type': 'application/json',
              // Include other headers as required, e.g., authorization
          },
          body: JSON.stringify(subjectObject)
      });

      if (!response.ok) {
          throw new Error(`HTTP error! status: ${response.status}`);
      }

      const data = await response.json();
      console.log('Data synced successfully:', data);
      // Handle success response, e.g., displaying a confirmation message to the user
  } catch (error) {
      console.error('Error syncing data:', error);
      // Handle errors, e.g., displaying an error message to the user
  }


  
  
}


function populateActions(actions, recipients) {
  if(actions && actions.length > 0){
    document.getElementById("action-list").style["display"] = "inline-block";

    const actionsTable = document.getElementById('actionsTable').getElementsByTagName('tbody')[0];

  actions.forEach((action, index) => {
    let row = actionsTable.insertRow(index);
    let actionCell = row.insertCell(0);
    let actionTextBox = document.createElement("input");
    actionTextBox.type = "text";
    actionTextBox.id = `action-${index}`;
    actionTextBox.value = action;
    actionTextBox.className = '.ms-textbox';

    let assignCell = row.insertCell(1);
    let selectList = document.createElement("select");
    selectList.id = `assignee-${index}`;
    selectList.className = 'ms-Dropdown';

    // Add an empty optiona
    let defaultOption = document.createElement("option");
    defaultOption.value = "";
    defaultOption.text = "Select recipient";
    selectList.appendChild(defaultOption);

    // Populate dropdown with recipients
    recipients.forEach(recipient => {
      let option = document.createElement("option");
      option.value = recipient;
      option.text = recipient;
      selectList.appendChild(option);
    });

    actionCell.appendChild(actionTextBox);
    assignCell.appendChild(selectList);
  });
  }
  
}
