# Standard library imports
import os
import sys

# Related third-party imports
import pandas as pd

try:
    import openpyxl
except ImportError:
    os.system('pip install openpyxl')

# Local application/library-specific imports
# import ace_tools as tools

# Creating a list of 10 topics and their content
data_10_entries = {
    "Week": [
        "Shared Channels", 
        "Using Planner in Teams", 
        "Meeting Notes Integration", 
        "Scheduling Assistant", 
        "Private vs. Public Teams", 
        "Tags and Mentions", 
        "Custom Backgrounds in Meetings", 
        "Using Breakout Rooms", 
        "Polls and Surveys", 
        "Task Management with To Do"
    ],
    "Tip of the Week": [
        "Shared Channels – Simplifying collaboration across departments",
        "Using Planner in Teams – Organize tasks and set due dates in a dedicated space",
        "Meeting Notes Integration – Capture meeting minutes and action items in one click",
        "Scheduling Assistant – Find the perfect time for meetings across time zones",
        "Private vs. Public Teams – How to organize your collaboration spaces effectively",
        "Tags and Mentions – Keep communications targeted and relevant",
        "Custom Backgrounds in Meetings – Personalize your video meetings for better engagement",
        "Using Breakout Rooms – Enhance collaboration in larger meetings with separate discussion spaces",
        "Polls and Surveys – Quickly gather feedback from your team during meetings",
        "Task Management with To Do – How to keep track of individual tasks in Teams"
    ],
    "What is [Feature Name]?": [
        "Shared Channels allow you to collaborate with people from other teams or departments without switching workspaces. These channels appear alongside your existing channels for seamless collaboration.",
        "Planner is a task management tool within Teams that lets you organize your team’s tasks in one shared space, assign tasks, set due dates, and track progress.",
        "Meeting Notes Integration lets you capture and share meeting minutes within Teams during your meeting. It helps ensure that everyone stays on the same page regarding action items and decisions.",
        "Scheduling Assistant helps you find the best time for meetings across time zones and between busy schedules. It eliminates the hassle of manual scheduling by showing everyone’s availability.",
        "Private vs. Public Teams allows you to control who can access the content shared within your team. Private Teams are great for sensitive projects, while Public Teams can be used for open collaboration.",
        "Tags and Mentions in Teams help you communicate more efficiently by grouping people into categories. Mention tags or individual names to make sure the right people get notified.",
        "Custom Backgrounds let you replace your video call background with an image of your choice, adding a personalized or professional look to your video meetings.",
        "Breakout Rooms allow you to split a large meeting into smaller discussion groups, enabling more focused collaboration during your sessions.",
        "Polls and Surveys in Teams allow you to collect instant feedback from your team members during meetings or within channels, ensuring everyone has a voice.",
        "Microsoft To Do is an integrated task management tool in Teams that allows you to track your individual tasks and stay organized."
    ],
    "How can you use it?": [
        "• Collaborating across departments is easier than ever. With Shared Channels, you can invite members from other teams into a channel without creating a whole new team.\n• Shared Channels also help ensure that conversations and files are centralized in one place.",
        "• Use Planner to break projects into manageable tasks. Assign responsibilities to team members, add deadlines, and track progress directly in Teams.\n• Create a new Planner from within a channel to integrate it with your ongoing discussions.",
        "• Start taking notes directly during your meeting in Teams, and share them with all participants afterward.\n• You can assign tasks from the meeting minutes so that action items don't fall through the cracks.",
        "• Use Scheduling Assistant to avoid back-and-forth email threads and schedule meetings efficiently.\n• The tool syncs with Outlook to show available times for all attendees and suggests the best options.",
        "• Use Private Teams for sensitive projects where you want to control access.\n• Public Teams are great for cross-department collaboration where anyone in the organization can join and contribute.",
        "• Use @tags to notify specific groups, like @marketing or @design, instead of mentioning individuals one by one.\n• Mentions make communication targeted and ensure that relevant people are notified.",
        "• Before a meeting, go to the ‘Background effects’ in your video settings and upload your chosen image.\n• This adds a professional or fun touch to your video calls, depending on the situation.",
        "• During a large meeting, select ‘Breakout rooms’ from the control panel.\n• Assign participants to rooms or let Teams automatically distribute them, making it easier to focus on specific discussions.",
        "• In your Teams chat or meeting, select the ‘Polls’ option to create a quick survey.\n• Gather responses in real-time and make data-driven decisions during the meeting.",
        "• Use the To Do app in Teams to manage your individual tasks.\n• Add new tasks, set deadlines, and prioritize your workload directly from Teams."
    ],
    "How to Get Started": [
        "1. In Teams, go to your team and create a Shared Channel.\n2. Invite members from other teams by entering their email addresses or selecting them from the directory.\n3. Start collaborating just like in regular channels.",
        "1. Go to the channel where you want to add a Planner.\n2. Click the “+” button at the top and select “Planner.”\n3. Start creating tasks and assigning them to your team members.",
        "1. During your Teams meeting, click on the “Meeting Notes” tab.\n2. Start taking notes and assigning tasks during the meeting.\n3. Share the notes and tasks with all participants after the meeting.",
        "1. Open the Calendar view in Teams or Outlook.\n2. Click on “New Meeting” and select “Scheduling Assistant.”\n3. Find the best time when all attendees are available and send out the invite.",
        "1. Go to Teams and create a new team.\n2. Select whether you want it to be a Private or Public team.\n3. Add your members and start collaborating.",
        "1. Create tags for your team by clicking on the ellipsis (…) next to the team name and selecting “Manage tags.”\n2. Use @mentions to notify specific people or groups in conversations.\n3. Ensure that important messages reach the right people.",
        "1. Open Teams, go to your video settings and select ‘Background effects.’\n2. Upload or choose an image from the gallery.\n3. Apply the background to your video meetings for a more polished appearance.",
        "1. During your Teams meeting, click on the Breakout Rooms option.\n2. Assign participants or let Teams do it automatically.\n3. Start the rooms and manage discussions in each smaller group.",
        "1. Click on the Polls tab during a meeting or in a channel.\n2. Create your questions and possible answers.\n3. Publish the poll and gather responses from your team.",
        "1. Open the To Do app in Teams from the left-hand sidebar.\n2. Start adding tasks and due dates to manage your workload.\n3. Track your progress as you complete tasks."
    ]
}

# Creating a DataFrame from the data
df_10_entries = pd.DataFrame(data_10_entries)

# Saving to Excel
output_file_10 = "./Microsoft_Teams_10_Friday_Tips.xlsx"
df_10_entries.to_excel(output_file_10, index=False)

# Display the file to the user

#tools.display_dataframe_to_user(name="Microsoft Teams 10 Friday Tips", dataframe=df_10_entries)

output_file_10