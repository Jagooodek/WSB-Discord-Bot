import discord
from discord.ext import commands
import pandas as pd
import re
from dotenv import load_dotenv
import os

load_dotenv()

intents = discord.Intents.default()
intents.members = True
intents.message_content = True
bot = commands.Bot(command_prefix='!', intents=intents)


async def verify():
    guild = bot.get_guild(int(os.environ.get('GUILD_ID')))

    log = open('error.log', 'w')
    current_log = 'Verifing started.\n'

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
            current_log += f'ERROR: wrong email, row = {index}, email = {email}\n'
            continue

        if discord_id not in members:
            discord_id = discord_id.lower()

        if discord_id not in members:
            log.write(f'ERROR: {discord_id} not found, row = {index}\n')
            current_log += f'ERROR: {discord_id} not found, row = {index}\n'
            continue

        member = members[discord_id]
        verified_members.add(member)
        roles_to_check = []
        roles_to_remove = []

        roles_to_check.append(roles["Verified"])


        if study_type == "Inżynierskich":
            roles_to_check.append(roles['Inżynier'])
            roles_to_check.append(roles[f'inż {group}'])

        elif study_type == "Licencjackich":
            roles_to_check.append(roles['Licencjat'])
            roles_to_check.append(roles[f'lic {group}'])

        roles_to_remove = [role for role in [*managed_roles, roles['Unverified']] if role not in roles_to_check and role in member.roles]
        roles_to_add = [role for role in roles_to_check if role not in member.roles]

        if len(roles_to_add) > 0:
            await member.add_roles(*roles_to_add)

        if len(roles_to_remove) > 0:
            await member.remove_roles(*roles_to_remove)
        if len(roles_to_remove) > 0 or len(roles_to_add) > 0:
            current_log += f'{member.name}: added = {[role.name for role in roles_to_add]}, removed = {[role.name for role in roles_to_remove]}\n'


    for member in members.values():
        if member not in verified_members:
            roles_to_remove = [role for role in managed_roles if role in member.roles]

            if len(roles_to_remove) > 0:
                await member.remove_roles(*roles_to_remove)
                current_log += f'{member.name}: removed = {[role.name for role in roles_to_remove]}\n'

            if roles['Unverified'] not in member.roles:
                await member.add_roles(roles['Unverified'])
                current_log += f'{member.name}: added = {roles["Unverified"].name}\n'
    current_log += "Done."
    return current_log

@bot.event
async def on_ready():
    print('Bot is ready')
    print(await verify())

@bot.event
async def on_member_join(member):
    print(await verify())

@bot.command(name="update_verified", description="This will update verified users")
async def update_verified(interaction: discord.Interaction, arg1: discord.Attachment):
    try:
        admin_role = discord.utils.get(interaction.guild.roles, name='Admin')
        if admin_role not in interaction.author.roles:
            await interaction.reply("Musisz być Adminem mordeczko")
            return
        await interaction.reply("Working...")
        await arg1.save(fp="source.xlsx")
        await interaction.reply(await verify())
    except Exception as e:
        await interaction.reply(str(e))

bot.run(os.environ.get('DISCORD_API_KEY'))
