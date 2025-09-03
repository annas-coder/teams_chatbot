require("dotenv").config();
const {
  TeamsActivityHandler,
  BotFrameworkAdapter,
  TurnContext,
  CardFactory,
} = require("botbuilder");
const express = require("express");
const mysql = require("mysql2");
const cron = require("node-cron");
const { v4: uuidv4 } = require("uuid"); // To generate unique IDs for form fields

const app = express();
app.use(express.json());

const port = process.env.PORT || 4000;

// MySQL connection pool
const pool = mysql.createPool({
  host: "localhost",
  user: "root",
  password: "password",
  database: "attendance",
  waitForConnections: true,
  connectionLimit: 10,
  queueLimit: 0,
});
// const promisePool = pool.promise();

// Store conversation references for proactive messages
const conversationReferences = {};

// Bot Adapter
const adapter = new BotFrameworkAdapter({
  appId: process.env.MICROSOFT_APP_ID,
  appPassword: process.env.MICROSOFT_APP_PASSWORD,
});

// Handle errors
adapter.onTurnError = async (context, error) => {
  console.error(`\n [onTurnError] unhandled error: ${error}`);
  await context.sendActivity("The bot encountered an error or bug.");
  await context.sendActivity(
    "To continue to run this bot, please fix the bot source code."
  );
};

// Teams Bot
class AttendanceBot extends TeamsActivityHandler {
  constructor() {
    super();

    // Store conversation reference for proactive messages
    this.onConversationUpdate(async (context, next) => {
      const conversationReference = TurnContext.getConversationReference(
        context.activity
      );
      conversationReferences[context.activity.from.id] = conversationReference;
      await next();
    });

    // Handle when a user sends a message
    this.onMessage(async (context, next) => {
      const text = (context.activity.text || "").trim().toLowerCase();

      if (text === "submit timesheet" || text === "timesheet") {
        const card = createTimesheetCard();
        await context.sendActivity({ attachments: [card] });
      } else if (text === "help") {
        await context.sendActivity(
          'Type "submit timesheet" to log your hours for the day.'
        );
      } else {
        await context.sendActivity(
          'Welcome to the Attendance Bot! Type "submit timesheet" to begin, or "help" for options.'
        );
      }
      await next();
    });

    //    this.onAdaptiveCardInvoke(async (context, next) => {
    //     const cardData = context.activity.value;
    //     const action = cardData.action;

    //     if (action === "AddProject") {
    //         await handleAddProject(context, cardData);
    //     } else if (action === "DeleteProject") {
    //         const projectIdToDelete = cardData.projectId;
    //         const currentProjects = cardData.projectData || [];

    //         const updatedProjects = currentProjects.filter(p => p.id !== projectIdToDelete);

    //         const newCard = createTimesheetCard(updatedProjects);
    //         await context.updateActivity({
    //             type: 'message',
    //             id: context.activity.replyToId,
    //             attachments: [newCard]
    //         });
    //     } else if (action === "SubmitTimesheet") {
    //         const userId = context.activity.from.id;
    //         await processTimesheetSubmission(cardData, userId, context);
    //     }

    //     await next();
    // });

    // Handle ALL Adaptive Card actions (Submit AND Add Project)
    // this.onAdaptiveCardInvoke(async (context, next) => {
    // const cardData = context.activity.value;
    // const userId = context.activity.from.id;

    // if (cardData.action === 'SubmitTimesheet') {
    //     await processTimesheetSubmission(cardData, userId, context);
    // } else if (cardData.action === 'AddProject') {
    //     // This handles the "Add Another Project" button click
    //     await handleAddProject(context, cardData);
    // }
    // await next();
    // });
  }
}

const bot = new AttendanceBot();

// Endpoint for Teams
app.post("/api/messages", (req, res) => {
  adapter.processActivity(req, res, async (context) => {
    await bot.run(context);
  });
});

// Function to create the initial Adaptive Card
function createTimesheetCard(projectData = []) {
  // If no existing data, start with one empty project
  if (projectData.length === 0) {
    projectData = [
      {
        id: uuidv4(), // Unique ID for this project block
        projectName: "",
        hours: 0,
        taskType: "development",
        tasks: "",
      },
    ];
  }

  const cardBody = buildCardBody(projectData);

  const cardJson = {
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.5",
    body: cardBody,
    actions: buildCardActions(projectData),
  };

  return CardFactory.adaptiveCard(cardJson);
}

// Build the main body of the card
function buildCardBody(projectData) {
  const body = [
    {
      type: "TextBlock",
      size: "Medium",
      weight: "Bolder",
      text: "Daily Timesheet",
      wrap: true,
    },
    {
      type: "TextBlock",
      text: "Please submit your timesheet for today. You can add multiple projects.",
      wrap: true,
      spacing: "Small",
    },
    {
      type: "Input.Date",
      id: "date",
      label: "Date",
      value: new Date().toISOString().split("T")[0],
    },
  ];

  // Add each project section
  projectData.forEach((project, index) => {
    body.push(...createProjectSection(project, index));
  });

  // Add total hours display
  const totalHours = projectData.reduce(
    (sum, project) => sum + (parseFloat(project.hours) || 0),
    0
  );
  body.push({
    type: "TextBlock",
    text: `Total Hours: ${totalHours.toFixed(1)}`,
    id: "totalHoursDisplay",
    weight: "Bolder",
    size: "Medium",
    spacing: "Medium",
  });

  // Add remarks
  body.push({
    type: "Input.Text",
    id: "remarks",
    label: "Remarks (Optional)",
    placeholder: "Any additional notes for the day...",
    isMultiline: true,
    spacing: "Medium",
  });

  return body;
}

// Create a section for a single project
// function createProjectSection(project, index) {
//     return [
//         {
//             "type": "Container",
//             "style": "emphasis",
//             "spacing": "Medium",
//             "items": [
//                 {
//                     "type": "TextBlock",
//                     "text": `Project #${index + 1}`,
//                     "weight": "Bolder",
//                     "size": "Medium",
//                     "wrap": true
//                 },
//                 {
//                     "type": "Input.ChoiceSet",
//                     "choices": [
//                         { "title": "Project Alpha", "value": "project-alpha" },
//                         { "title": "Project Beta", "value": "project-beta" },
//                         { "title": "Project Gamma", "value": "project-gamma" },
//                         { "title": "Internal / Admin", "value": "internal" }
//                     ],
//                     "placeholder": "Select a Project",
//                     "id": `project-${project.id}`,
//                     "label": "Project Name *",
//                     "value": project.projectName
//                 },
//                 {
//                     "type": "Input.Number",
//                     "id": `hours-${project.id}`,
//                     "label": "Hours Worked *",
//                     "min": 0,
//                     "max": 24,
//                     "value": project.hours
//                 },
//                 {
//                     "type": "Input.ChoiceSet",
//                     "choices": [
//                         { "title": "Development", "value": "development" },
//                         { "title": "Design", "value": "design" },
//                         { "title": "Meeting", "value": "meeting" },
//                         { "title": "Testing", "value": "testing" },
//                         { "title": "Research", "value": "research" },
//                         { "title": "Documentation", "value": "documentation" }
//                     ],
//                     "placeholder": "Select Task Type",
//                     "id": `task-type-${project.id}`,
//                     "label": "Task Type",
//                     "value": project.taskType
//                 },
//                 {
//                     "type": "Input.Text",
//                     "id": `tasks-${project.id}`,
//                     "label": "Tasks Completed *",
//                     "placeholder": "Describe what you worked on for this project...",
//                     "isMultiline": true,
//                     "value": project.tasks
//                 }
//             ]
//         }
//     ];
// }

function createProjectSection(project, index) {
  return [
    {
      type: "Container",
      style: "emphasis",
      spacing: "Medium",
      items: [
        {
          type: "TextBlock",
          text: `Project #${index + 1}`,
          weight: "Bolder",
          size: "Medium",
          wrap: true,
        },
        {
          type: "Input.ChoiceSet",
          choices: [
            { title: "Project Alpha", value: "project-alpha" },
            { title: "Project Beta", value: "project-beta" },
          ],
          placeholder: "Select a Project",
          id: `project-${project.id}`,
          label: "Project Name *",
          value: project.projectName,
        },
        {
          type: "Input.Number",
          id: `hours-${project.id}`,
          label: "Hours Worked *",
          min: 0,
          max: 24,
          value: project.hours,
        },
        {
          type: "Input.Text",
          id: `tasks-${project.id}`,
          label: "Tasks Completed *",
          isMultiline: true,
          value: project.tasks,
        },
        {
          type: "Action.Submit",
          title: "Delete Project",
          data: {
            action: "DeleteProject",
            projectId: project.id,
          },
        },
      ],
    },
  ];
}

// Build the action buttons at the bottom of the card
function buildCardActions(projectData) {
  return [
    {
      type: "Action.Submit",
      title: "Submit Timesheet",
      data: { action: "SubmitTimesheet" },
    },
    {
      type: "Action.Submit",
      title: "Submit Timesheet",
      data: { action: "SubmitTimesheet" },
    },
  ];
}

// Handle the "Add Another Project" button click
async function handleAddProject(context, cardData) {
  // Extract current project data from the card
  const currentProjectData = cardData.projectData || [];

  // Add a new empty project
  currentProjectData.push({
    id: uuidv4(),
    projectName: "",
    hours: 0,
    taskType: "development",
    tasks: "",
  });

  // Create a new card with the additional project
  const newCard = createTimesheetCard(currentProjectData);

  // Update the existing message with the new card
  await context.updateActivity({
    type: "message",
    id: context.activity.replyToId,
    attachments: [newCard],
  });
}

// Process the final timesheet submission
async function processTimesheetSubmission(cardData, userId, context) {
  const submissionDate = cardData.date;
  const remarks = cardData.remarks || "";
  const projectData = cardData.projectData || [];
  let totalHours = 0;

  try {
    // Process each project
    for (const project of projectData) {
      const projectFieldId = project.id;
      const projectId = cardData[`project-${projectFieldId}`];
      const hours = parseFloat(cardData[`hours-${projectFieldId}`]);
      const taskType = cardData[`task-type-${projectFieldId}`];
      const tasksCompleted = cardData[`tasks-${projectFieldId}`];

      // Validate required fields
      if (!projectId || isNaN(hours) || !tasksCompleted) {
        await context.sendActivity(
          "❌ Error: Please fill in all required fields for all projects (Project, Hours, Tasks)."
        );
        return;
      }

      totalHours += hours;

      // Insert each project into the database
      await promisePool.query(
        `INSERT INTO submissions (employee_id, project_id, date, hours, task_type, tasks, remarks, timestamp) 
                 VALUES (?, ?, ?, ?, ?, ?, ?, NOW())`,
        [
          userId,
          projectId,
          submissionDate,
          hours,
          taskType,
          tasksCompleted,
          remarks,
        ]
      );
    }

    // Send success message
    if (totalHours < 8) {
      await context.sendActivity(
        `⚠ Warning: Total hours (${totalHours}) less than 8! Submission was still recorded.`
      );
    } else {
      await context.sendActivity(
        `✅ Timesheet submitted successfully! Total hours: ${totalHours}`
      );
    }
  } catch (error) {
    console.error("Database insertion error:", error);
    await context.sendActivity(
      "❌ Sorry, there was an error saving your timesheet. Please try again."
    );
  }
}

// Modified Cron Job to send proactive reminder with card
cron.schedule("0 18 * * *", async () => {
  console.log("Running 6 PM reminder job...");
  try {
    const [employees] = await promisePool.query(
      "SELECT teams_id, name FROM employees"
    );

    for (const emp of employees) {
      const conversationReference = conversationReferences[emp.teams_id];
      if (conversationReference) {
        try {
          await adapter.continueConversation(
            conversationReference,
            async (turnContext) => {
              await turnContext.sendActivity(
                `Hi ${emp.name}! 👋 Don't forget to submit your timesheet for today.`
              );
              const card = createTimesheetCard();
              await turnContext.sendActivity({ attachments: [card] });
            }
          );
          console.log(`Reminder and card sent to ${emp.name}`);
        } catch (proactiveError) {
          console.error(
            `Failed to send proactive message to ${emp.name}:`,
            proactiveError
          );
        }
      } else {
        console.log(
          `No conversation reference found for ${emp.name}. They need to message the bot first.`
        );
      }
    }
  } catch (dbError) {
    console.error("Database error in cron job:", dbError);
  }
});

// Health check endpoint
app.get("/testing", (req, res) => {
  res.send(`
        <!DOCTYPE html>
        <html>
        <head>
            <title>Test Page</title>
        </head>
        <body>
            <h1>Hello from Attendance Bot!</h1>
            <p>This bot now supports multiple projects with add/remove functionality.</p>
        </body>
        </html>
    `);
});

app.listen(port, () => console.log("Bot running on port 3978"));
