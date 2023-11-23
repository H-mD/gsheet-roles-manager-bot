import discord
from discord.ext import commands
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from time import sleep
from dotenv import load_dotenv
import os

load_dotenv()

# Bot 
TOKEN = os.environ.get("TOKEN")
intents = discord.Intents.default()
intents.message_content = True
intents.members = True
intents.guilds = True
bot = commands.Bot(command_prefix=';', intents=intents)

# GSheet
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(os.environ.get("CREDENTIALS_FILE"), scope)
client = gspread.authorize(creds)

def switch_case(string):
    if string[0].islower():
        return string[0].upper() + string[1:]
    elif string[0].isupper():
        return string[0].lower() + string[1:]
    else:
        return string

# Bot Commands
@bot.command()
async def removeRole(ctx, member_name, role_name):
    guild = ctx.guild
    role = discord.utils.get(guild.roles, name=role_name)
    member = discord.utils.get(guild.members, name=member_name)
    
    try:
        if member.guild != ctx.guild:
            await ctx.send("Member not found in this server.")
        elif role.guild != ctx.guild:
            await ctx.send("Role not found in this server.")
        else:
            await member.remove_roles(role)
            await ctx.send(f"Removed {role.name} from {member.display_name}")
    except discord.Forbidden:
        await ctx.send("I don't have permission to do that.")
    except discord.HTTPException:
        await ctx.send("Failed to add role. An error occurred.")

@bot.command()
async def addRole(ctx, member_name, role_name):
    guild = ctx.guild
    role = discord.utils.get(guild.roles, name=role_name)
    member = discord.utils.get(guild.members, name=member_name)
    
    try:
        if member.guild != ctx.guild:
            await ctx.send("Member not found in this server.")
        elif role.guild != ctx.guild:
            await ctx.send("Role not found in this server.")
        else:
            await member.add_roles(role)
            await ctx.send(f"Added {role.name} to {member.display_name}")
    except discord.Forbidden:
        await ctx.send("I don't have permission to do that.")
    except discord.HTTPException:
        await ctx.send("Failed to add role. An error occurred.")

@bot.command()
async def batchRole(ctx, role_name=None, sheet_url=None, sheet_number: int=None, header=None):
    guild = ctx.guild
    role = discord.utils.get(guild.roles, name=role_name)

    # embed = discord.Embed(
    #     title="bothehe",
    #     description="bot specifically made for AJK Lab's event utilities",
    # )
    # embed.add_field(name="Process", value="Value 1", inline=False)
    # embed.add_field(name="Progress", value="Value 2", inline=False)

    # await ctx.send(embed=embed)
    
    if role_name is None:
        await ctx.send(f"command format: ;batchRole <role_name> <sheet_url> <sheet_number> <header>")
    elif sheet_url is None:
        await ctx.send(f"command format: ;batchRole <role_name> <sheet_url> <sheet_number> <header>")
    elif sheet_number is None:
        await ctx.send(f"command format: ;batchRole <role_name> <sheet_url> <sheet_number> <header>")
    elif header is None:
        await ctx.send(f"command format: ;batchRole <role_name> <sheet_url> <sheet_number> <header>")
    elif role.guild != ctx.guild:
        await ctx.send(f"Role **[{role_name}]** not found in this server.")
    else:
        try:
            sheet = client.open_by_url(sheet_url)
            worksheet = sheet.get_worksheet(sheet_number)

            data_header = worksheet.find(header)
            status_header = worksheet.find("status")
            angkatan_header = worksheet.find("angkatan")

            if data_header is None:
                await ctx.send(f"header **[{header}]** not found")
            else:
                data_range = worksheet.range(data_header.row + 1, data_header.col, worksheet.row_count, data_header.col)
                angkatan_range = worksheet.range(angkatan_header.row + 1, angkatan_header.col, worksheet.row_count, angkatan_header.col)
                data_len = 0
                for i in data_range:
                    if i.value == "":
                        break
                    data_len += 1

                for i, data in enumerate(data_range):
                    sleep(1)
                    if data.value == "":
                        await ctx.send(f"***Done***")
                        break
                    if data.value == header:
                        continue

                    # filter member
                    if angkatan_range[i].value == "2021":
                        worksheet.update_cell(data_header.row+1+i, status_header.col, "x")
                        # await ctx.send(f"({data_len}/{i+1})  Skipping **[{data.value}]**, Reason: *Angkatan 2021*")
                        continue
                    elif angkatan_range[i].value == "2020":
                        worksheet.update_cell(data_header.row+1+i, status_header.col, "x")
                        # await ctx.send(f"({data_len}/{i+1})  Skipping **[{data.value}]**, Reason: *Angkatan 2020*")
                        continue
                    if worksheet.cell(data_header.row+1+i, status_header.col).value == "v":
                        # await ctx.send(f"({data_len}/{i+1})  Skipping **[{data.value}]**, Reason: *Already Added*")
                        continue
                    
                    # find member name
                    member_name = data.value.split("#")[0] # remove discriminator
                    member_name = member_name.lstrip("@") # remove @
                    member_name = member_name.rstrip() # remove trailing space
                    member = discord.utils.get(guild.members, name=member_name) # search username
                    if member is None:
                        member = discord.utils.get(guild.members, nick=member_name) # search nickname
                        if member is None:
                            member_name = switch_case(member_name) # try switch case
                            member = discord.utils.get(guild.members, name=member_name) # search username
                            if member is None:
                                member = discord.utils.get(guild.members, nick=member_name) # search nickname
                                if member is None:
                                    worksheet.update_cell(data_header.row+1+i, status_header.col, "x")
                                    # await ctx.send(f"({data_len}/{i+1})  Skipping **[{member_name}]**, Reason: *Member not found*")
                                    continue
                    # add role
                    await member.add_roles(role)
                    await ctx.send(f"({data_len}/{i+1})  Added *{role_name}* to **[{member}]**")
                    worksheet.update_cell(data_header.row+1+i, status_header.col, "v")

        except discord.Forbidden:
            await ctx.send(f"I don't have permission to do that.")
        except discord.HTTPException:
            await ctx.send(f"Failed to add role. An error occurred.")
        except gspread.exceptions.SpreadsheetNotFound:
            await ctx.send(f"Spreadsheet not found")

@bot.command()
async def batchFix(ctx, role_name=None, sheet_url=None, sheet_number: int=None, header=None):
    guild = ctx.guild
    role = discord.utils.get(guild.roles, name=role_name)

    embed = discord.Embed(title="Fixing Role")
    embed.description = f"Processing ..."

    msg = await ctx.send(embed=embed)
    
    if role_name is None:
        embed.description = (f"command format: **;batchFix <role_name> <sheet_url> <sheet_number> <header>**")
        await msg.edit(embed=embed)
    elif sheet_url is None:
        embed.description = (f"command format: **;batchFix <role_name> <sheet_url> <sheet_number> <header>**")
        await msg.edit(embed=embed)
    elif sheet_number is None:
        embed.description = (f"command format: **;batchFix <role_name> <sheet_url> <sheet_number> <header>**")
        await msg.edit(embed=embed)
    elif header is None:
        embed.description = (f"command format: **;batchFix <role_name> <sheet_url> <sheet_number> <header>**")
        await msg.edit(embed=embed)
    elif role.guild != ctx.guild:
        embed.description = (f"ERROR\nRole **[{role_name}]** not found in this server.")
        await msg.edit(embed=embed)
    else:
        try:
            sheet = client.open_by_url(sheet_url)
            worksheet = sheet.get_worksheet(sheet_number)

            data_header = worksheet.find(header)
            status_header = worksheet.find("status")
            angkatan_header = worksheet.find("angkatan")

            if data_header is None:
                await ctx.send(f"header **[{header}]** not found")
            else:
                data_range = worksheet.range(data_header.row + 1, data_header.col, worksheet.row_count, data_header.col)
                angkatan_range = worksheet.range(angkatan_header.row + 1, angkatan_header.col, worksheet.row_count, angkatan_header.col)
                status_range = worksheet.range(status_header.row + 1, status_header.col, worksheet.row_count, status_header.col)
                data_len = 0
                for i in status_range:
                    if i.value == "":
                        break
                    data_len += 1

                for i, data in enumerate(data_range):
                    sleep(1)
                    if data.value == header:
                        continue

                    if i+1 > data_len:
                        embed.description += (f"\n**DONE**")
                        await msg.edit(embed=embed)
                        break

                    embed.description = (f"Progress: **{data_len}/{i+1}**")
                    await msg.edit(embed=embed)

                    if angkatan_range[i].value == "2021":
                        worksheet.update_cell(data_header.row+1+i, status_header.col, "x")
                        continue
                    elif angkatan_range[i].value == "2020":
                        worksheet.update_cell(data_header.row+1+i, status_header.col, "x")
                        continue
                    
                    # find member name
                    member_name = data.value.split("#")[0] # remove discriminator
                    member_name = member_name.lstrip("@") # remove @

                    member = discord.utils.get(guild.members, name=member_name) # search username
                    if member is None:
                        member = discord.utils.get(guild.members, nick=member_name) # search nickname
                        if member is None:
                            member_name = switch_case(member_name) # try switch case
                            member = discord.utils.get(guild.members, name=member_name) # search username
                            if member is None:
                                member = discord.utils.get(guild.members, nick=member_name) # search nickname
                                if member is None:
                                    worksheet.update_cell(data_header.row+1+i, status_header.col, "x")
                                    continue
                    
                    # check member role
                    if role not in member.roles:
                        await member.add_roles(role)
                    worksheet.update_cell(data_header.row+1+i, status_header.col, "v")

        except discord.Forbidden:
            embed.description = (f"ERROR\nI don't have permission to do that.")
            await msg.edit(embed=embed)
        except discord.HTTPException:
            embed.description = (f"ERROR\nFailed to add role. An error occurred.")
            await msg.edit(embed=embed)
        except gspread.exceptions.SpreadsheetNotFound:
            embed.description = (f"ERROR\nSpreadsheet not found")
            await msg.edit(embed=embed)

bot.run(TOKEN)