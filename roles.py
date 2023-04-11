import discord
import pandas as pd
from config import *

intents = discord.Intents.default()
intents.members = True
client = discord.Client(intents=intents)

@client.event
async def on_ready():
    try:
        print(f"Logged in as {client.user.name}")
    except Exception as e:
        print(f"Error in on_ready event handler: {e}")
        
    df = pd.read_excel(EXCEL_SHEET_PATH) 

    # Get a list of all the members in the server
    guild = discord.utils.get(client.guilds, name=SERVER_NAME)
    members = []
    async for member in guild.fetch_members(limit=None):
        members.append(member)

    # Loop through each name in the spreadsheet
    for _, row in df.iterrows():
        name = row.loc[ROW_NAME]
        for member in members:
            # Lower case the name and look if there's a match between discord and the excel spreadsheet
            if name.lower() == member.display_name.lower():
                # Change the role of the member
                role = discord.utils.get(guild.roles, name=ROLE_NAME)
                await member.add_roles(role)
                print(f"{member.display_name} has been given the {ROLE_NAME} role.")

client.run(DISCORD_API_TOKEN)