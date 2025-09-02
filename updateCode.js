require('dotenv').config();
const { TeamsActivityHandler, CardFactory, BotFrameworkAdapter, TurnContext } = require('botbuilder');
const express = require('express');
const { v4: uuidv4 } = require('uuid');

const app = express();
app.use(express.json());

const adapter = new BotFrameworkAdapter({
    appId: process.env.MICROSOFT_APP_ID,
    appPassword: process.env.MICROSOFT_APP_PASSWORD
});

// In-memory user sessions (per userId)
const userProjects = {};
let lastCardIds = {};

class TimesheetBot extends TeamsActivityHandler {
    constructor() {
        super();

        this.onMessage(async (context, next) => {
            const text = (context.activity.text || '').trim().toLowerCase();
            const userId = context.activity.from.id;

            const value = context.activity.value; // Adaptive Card submission

            if (value) {
                if (value.action === 'AddProject') {
                    await handleAddProject(context, userId, value);
                } else if (value.action === 'DeleteProject') {
                    await handleDeleteProject(context, userId, value);
                } else if (value.action === 'SubmitTimesheet') {
                    await handleSubmitTimesheet(context, userId, value);
                } 
                return;
            }

            if (text === 'timesheet') {
                userProjects[userId] = [{ id: uuidv4(), projectName: '', hours: 0 }];
                const card = createTimesheetCard(userProjects[userId]);
                const sentActivity = await context.sendActivity({ attachments: [card] });
                lastCardIds[context.activity.from.id] = sentActivity.id;
            } else {
                await context.sendActivity('Type "timesheet" to start filling your timesheet.'); 
            }

            await next(); 
        });
    }
}

let currentProjects = [];
// === Handlers ===
async function handleAddProject(context, userId, cardData) {
if (!userProjects[userId]) {
        userProjects[userId] = [];
    }

    // Add a new empty project
    userProjects[userId].push({
        id: uuidv4(),
        projectName: '',
        hours: 0
    });

    const newCard = createTimesheetCard(userProjects[userId]);
    const cardId = lastCardIds[userId];

    if (cardId) {
        await context.updateActivity({
            type: 'message',
            id: cardId,
            conversation: context.activity.conversation,
            attachments: [newCard]
        });
    } else {
        const sentActivity = await context.sendActivity({ attachments: [newCard] });
        lastCardIds[userId] = sentActivity.id;
    }

}

async function handleDeleteProject(context, userId, cardData) {
    userProjects[userId] = userProjects[userId].filter(p => p.id !== cardData.projectId);
    const newCard = createTimesheetCard(userProjects[userId]);
    await context.updateActivity({
        type: 'message',
        id: context.activity.replyToId,
        attachments: [newCard]
    });
}

async function handleSubmitTimesheet(context, userId, cardData) {
    // Extract values from submitted fields
    const submittedProjects = userProjects[userId].map(p => ({
        id: p.id,
        projectName: cardData[`project-${p.id}`],
        hours: cardData[`hours-${p.id}`]
    }));

    console.log("✅ Final Timesheet:", {
        date: cardData.date,
        projects: submittedProjects
    });

    await context.sendActivity(`✅ Timesheet submitted with ${submittedProjects.length} project(s). Thanks!`);

    // Clear session
    delete userProjects[userId];
}

// === Card Builder ===
function createTimesheetCard(projectData = []) {
    const body = [
        { type: 'TextBlock', text: 'Daily Timesheet', weight: 'Bolder', size: 'Medium' },
        { type: 'Input.Date', id: 'date', label: 'Date', value: new Date().toISOString().split('T')[0] }
    ];

    projectData.forEach((p, index) => {
        body.push(...createProjectRow(p, index));
    });

    return CardFactory.adaptiveCard({
        $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
        type: 'AdaptiveCard',
        version: '1.5',
        body,
        actions: [
            { type: 'Action.Submit', title: '➕ Add Project', data: { action: 'AddProject' } },
            { type: 'Action.Submit', title: '✅ Submit Timesheet', data: { action: 'SubmitTimesheet' } }
        ]
    });
}

function createProjectRow(project, index) {
    return [
        {
            type: 'Container',
            style: 'emphasis',
            spacing: 'Medium',
            items: [
                { type: 'TextBlock', text: `Project #${index + 1}`, weight: 'Bolder' },
                { type: 'Input.Text', id: `project-${project.id}`, label: 'Project Name', value: project.projectName },
                { type: 'Input.Number', id: `hours-${project.id}`, label: 'Hours Worked', value: project.hours },
                {
                    type: 'Action.Submit',
                    title: '🗑 Delete Project',
                    data: { action: 'DeleteProject', projectId: project.id }
                }
            ]
        }
    ];
}

// === Express Bot Endpoint ===
const bot = new TimesheetBot();
app.post('/api/messages', (req, res) => {
    adapter.processActivity(req, res, async (context) => {
        await bot.run(context);
    });
});

app.listen(3978, () => console.log('🚀 Bot running on port 3978'));
