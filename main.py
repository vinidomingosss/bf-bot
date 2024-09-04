import discord
import pandas as pd
from discord.ext import commands
from datetime import datetime, timedelta, timezone
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from apscheduler.triggers.cron import CronTrigger
import os
from dotenv import load_dotenv

intents = discord.Intents.default()
intents.message_content = True
intents.members = True

# Carrega as variáveis de ambiente do arquivo .env
load_dotenv()

# Configura o bot com o token do Discord
TOKEN = os.getenv('id_token')
bot = commands.Bot(command_prefix='!', intents=intents)

# Lista para armazenar mensagens
mensagens = []

# Variável para armazenar a data/hora da última geração do Excel
ultima_geracao = None

# ID do canal específico para o comando de gerar Excel
ID_CANAL_ESPECIFICO = 1279957234930942007  

# ID do cargo de suporte 
SUPORTE_ROLE_ID = 1280304161929298010

@bot.event
async def on_ready():
    global ultima_geracao
    print(f'Bot conectado como {bot.user}')
    # Define a data/hora atual como a última geração inicial, usando UTC
    ultima_geracao = datetime.now(timezone.utc)

    # Configura o agendador para a geração automática do Excel
    scheduler = AsyncIOScheduler()
    trigger_time = datetime.now(timezone.utc) + timedelta(minutes=2)
    scheduler.add_job(gerar_excel_automatico, CronTrigger(start_date=trigger_time))
    #scheduler.add_job(gerar_excel_automatico, CronTrigger(day_of_week='mon-fri', hour=21, minute=21))
    print("Scheduler configurado!")
    scheduler.start()
    print("Scheduler iniciado!")

@bot.event
async def on_message(message):
    if message.author == bot.user:
        return

    global ultima_geracao

    # Converte o conteúdo da mensagem para maiúsculas
    content_upper = message.content.upper()

    # Verifica se a mensagem é do canal específico
    if message.channel.id == ID_CANAL_ESPECIFICO:
        # Converte ambos os datetime para "offset-naive"
        message_time = message.created_at.astimezone(timezone.utc).replace(tzinfo=None)
        ultima_geracao_naive = ultima_geracao.replace(tzinfo=None)

        # Verifica se a mensagem foi enviada após a última geração do Excel
        if message_time > ultima_geracao_naive:
            # Verifica se a mensagem começa com "1171" ou "PBA" (em maiúsculas)
            if content_upper.startswith("1171") or content_upper.startswith("PBA"):
                # Armazena a mensagem em uma lista
                mensagens.append({
                'Numero terminal': message.content.upper()
                })

    # Verifique comandos ou responda ao usuário
    await bot.process_commands(message)

@bot.command(name='gerar')
async def gerar(ctx):
    global ultima_geracao

    # Verifica se o comando está sendo executado no canal específico
    if ctx.channel.id != ID_CANAL_ESPECIFICO:
        await ctx.channel.send("Este comando só pode ser executado no canal do suporte!")
        return

    # Cria um DataFrame a partir das mensagens filtradas
    df = pd.DataFrame(mensagens)
    
    # Salva o DataFrame em um arquivo Excel
    df.to_excel('terminais.xlsx', index=False, header=False)
    
    # Envia o arquivo Excel no canal
    await ctx.channel.send(file=discord.File('terminais.xlsx'))

    # Atualiza a data/hora da última geração do Excel usando UTC
    ultima_geracao = datetime.now(timezone.utc)

    # Limpa a lista de mensagens após gerar o Excel
    mensagens.clear()

@bot.event
async def on_command_error(ctx, error):
    if isinstance(error, commands.CommandNotFound):
        # Mensagem de comando inválido
        await ctx.send("Comando inválido. Aqui está a lista de comandos disponíveis:")
        
        # Gera a lista de comandos disponíveis
        commands_list = "\n".join([f"!{command.name}" for command in bot.commands if not command.hidden])
        
        # Envia a lista de comandos
        await ctx.send(f"Comandos disponíveis:\n{commands_list}")        

@bot.command()
async def ola(ctx):
    await ctx.send('Olá, Mensagem para testar o BOT!')

async def gerar_excel_automatico():

    print("Gerando Excel automaticamente...")
    # Obtém o canal específico
    channel = bot.get_channel(ID_CANAL_ESPECIFICO)
    if channel:

        print(f"Canal encontrado: {channel.name}")
        # Cria um contexto de comando falso
        class FakeContext:
            def __init__(self, channel):
                self.channel = channel
            async def send(self, *args, **kwargs):
                return await channel.send(*args, **kwargs)

        ctx = FakeContext(channel)
        await gerar(ctx)
        
        # Marca o cargo "suporte"
        suporte_role = discord.utils.get(channel.guild.roles, id=SUPORTE_ROLE_ID)
        if suporte_role:
            await channel.send(f"✉️ Não esqueçam de enviar os terminais pessoal! {suporte_role.mention}")
        else:
            await channel.send("O Excel foi gerado automaticamente, mas não foi possível marcar o cargo suporte.")

bot.run(TOKEN)