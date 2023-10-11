import discord
from discord.ext import commands
import pandas as pd
import re
from dotenv import load_dotenv
import os

load_dotenv()

intents = discord.Intents.default()
intents.members = True
bot = commands.Bot(command_prefix='!', intents=intents)

@bot.event
async def on_ready():
    print('Bot is ready')
    guild = bot.get_guild(int(os.environ.get('GUILD_ID')))

    log = open('error.log', 'w')

    members = {}
    for member in guild.members:
        members[member.name] = member

    roles = {}
    for role in guild.roles:
        roles[role.name] = role

    managed_roles = [roles['Inżynier'], roles['Licencjat'], roles['Verified']]
    managed_roles.extend([roles[f'inż {group}'] for group in range(1, 7)])
    managed_roles.extend([roles[f'lic {group}'] for group in range(1, 7)])

    df = pd.read_excel("source.xlsx")
    verified_members = set()
    for index, row in df.iterrows():

        discord_id = row.iloc[-2]
        discord_id = re.sub(r'#\d+$', '', discord_id)
        study_type = row.iloc[-3]
        group = row.iloc[-1]
        email = row.iloc[3]

        while discord_id[-1].isspace():
            discord_id = discord_id[:-1]
        while discord_id[0].isspace():
            discord_id = discord_id[1:]

        if not email.endswith('@student.wroclaw.merito.pl'):
            log.write(f'ERROR: wrong email, row = {index}, email = {email}\n')
            print(f'ERROR: wrong email = {email}')
            continue

        if discord_id not in members:
            discord_id = discord_id.lower()

        if discord_id not in members:
            print(f'ERROR: {discord_id} not found')
            log.write(f'ERROR: {discord_id} not found, row = {index}\n')
            continue

        member = members[discord_id]
        verified_members.add(member)
        roles_to_check = []
        roles_to_remove = []

        roles_to_check.append(roles["Verified"])
        roles_to_remove.append(roles["Unverified"])

        if study_type == "Inżynierskich":
            roles_to_check.append(roles['Inżynier'])
            roles_to_check.append(roles[f'inż {group}'])

        elif study_type == "Licencjackich":
            roles_to_check.append(roles['Licencjat'])
            roles_to_check.append(roles[f'lic {group}'])

        roles_to_remove = [role for role in managed_roles if role not in roles_to_check and role in member.roles]
        roles_to_add = [role for role in roles_to_check if role not in member.roles]

        if len(roles_to_add) > 0:
            await member.add_roles(*roles_to_add)

        if len(roles_to_remove) > 0:
            await member.remove_roles(*roles_to_remove)
        if len(roles_to_remove) > 0 or len(roles_to_add) > 0:
            print(f'{member.name}: added = {[role.name for role in roles_to_add]}, removed = {[role.name for role in roles_to_remove]}')


    for member in members.values():
        if member not in verified_members:
            roles_to_remove = [role for role in managed_roles if role in member.roles]

            if len(roles_to_remove) > 0:
                await member.remove_roles(*roles_to_remove)
                print(f'{member.name}: removed = {[role.name for role in roles_to_remove]}')

            if roles['Unverified'] not in member.roles:
                await member.add_roles(roles['Unverified'])
                print(f'{member.name}: added = {roles["Unverified"].name} ')
    print("DONE")


bot.run(os.environ.get('DISCORD_API_KEY'))
