import streamlit as st
from streamlit_option_menu import option_menu
import pandas as pd
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
from pytz import timezone
import matplotlib.pyplot as plt
import calendar
import time

# Definir timezone para Brasília
br_timezone = timezone('America/Sao_Paulo')

# Função para enviar e-mail
def send_email(subject, body, to_email, cc_emails=[], image_path=None):
    from_email = "manutencaomnh@gmail.com"
    password = "axhb fsec ezcg txbz"
    
    # Configurar a mensagem do e-mail
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject
    if cc_emails:
        msg['Cc'] = ', '.join(cc_emails)
    msg.attach(MIMEText(body, 'html'))
    
    # Anexar imagem, se existir
    if image_path and os.path.exists(image_path):
        with open(image_path, "rb") as f:
            mime = MIMEBase('image', 'png', filename=os.path.basename(image_path))
            mime.add_header('Content-Disposition', 'attachment', filename=os.path.basename(image_path))
            mime.add_header('X-Attachment-Id', '0')
            mime.add_header('Content-ID', '<0>')
            mime.set_payload(f.read())
            encoders.encode_base64(mime)
            msg.attach(mime)
    
    # Enviar o e-mail
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(from_email, password)
        text = msg.as_string()
        recipients = [to_email] + cc_emails
        server.sendmail(from_email, recipients, text)
        server.quit()
        print("Email enviado com sucesso!")
    except Exception as e:
        print(f"Erro ao enviar email: {e}")

# Função para gerar número sequencial de OS
def generate_os_number():
    os_number_file = 'last_os_number.txt'
    if not os.path.exists(os_number_file):
        with open(os_number_file, 'w') as f:
            f.write('0')
    
    with open(os_number_file, 'r') as f:
        last_os_number = int(f.read())
    
    new_os_number = last_os_number + 1
    
    with open(os_number_file, 'w') as f:
        f.write(str(new_os_number))
    
    return new_os_number

def save_data(df, file_name):
    df.to_csv(file_name, index=False)

def load_data(file_name, default_list):
    try:
        df = pd.read_csv(file_name)
    except FileNotFoundError:
        df = pd.DataFrame(default_list, columns=['Item'])
        save_data(df, file_name)
    return df

def save_emails(emails):
    with open('emails.txt', 'w') as file:
        for email in emails:
            file.write(f"{email}\n")

def load_emails():
    try:
        with open('emails.txt', 'r') as file:
            emails = file.read().splitlines()
    except FileNotFoundError:
        emails = []
    return emails

def save_os(df, os_data, image_path, cc_emails):
    new_row = pd.DataFrame([os_data])
    df = pd.concat([df, new_row], ignore_index=True)
    save_data(df, 'os_data.csv')
    
    # Enviar e-mail de abertura de OS
    subject = f"OS {os_data['Numero_OS']} - {os_data['Equipamento_Setor']} - {os_data['Motivo_Parada']} - {os_data['Status']} - {os_data['Data_Hora']}"
    body = f"""
    <html>
    <body>
        <p>Detalhes da OS:</p>
        <p>Número: {os_data['Numero_OS']}</p>
        <p>Equipamento/Setor: {os_data['Equipamento_Setor']}</p>
        <p>Motivo: {os_data['Motivo_Parada']}</p>
        <p>Status: {os_data['Status']}</p>
        <p>Data e Hora: {os_data['Data_Hora']}</p>
        <br>
        <img src="cid:0" alt="Imagem anexada">
    </body>
    </html>
    """
    send_email(subject, body, "manutencaomnh@gmail.com", cc_emails, image_path)
    
    return df

def close_os(df, os_number, item_usado, manutencao_com, data_hora, cc_emails):
    df.loc[df['Numero_OS'] == os_number, ['Item_Usado', 'Manutencao_Com', 'Status', 'Data_Hora_Fechamento']] = [item_usado, manutencao_com, 'Fechada', data_hora]
    save_data(df, 'os_data.csv')
    
    # Enviar e-mail de fechamento de OS
    os_data = df[df['Numero_OS'] == os_number].iloc[0]
    subject = f"OS {os_data['Numero_OS']} - {os_data['Equipamento_Setor']} - {os_data['Motivo_Parada']} - Fechada - {data_hora}"
    body = f"""
    <html>
    <body>
        <p>Detalhes da OS Fechada:</p>
        <p>Número: {os_number}</p>
        <p>Item Usado: {item_usado}</p>
        <p>Manutenção Feita com: {manutencao_com}</p>
        <p>Status: Fechada</p>
        <p>Data e Hora: {data_hora}</p>
    </body>
    </html>
    """
    send_email(subject, body, "manutencaomnh@gmail.com", cc_emails)

# Carregar dados
df_os = load_data('os_data.csv', [{'Numero_OS': 0, 'Equipamento_Setor': '', 'Motivo_Parada': '', 'Imagem': '', 'Status': '', 'Data_Hora': '', 'Item_Usado': '', 'Manutencao_Com': '', 'Data_Hora_Fechamento': ''}])
equipamentos_setores = load_data('equipamentos_setores.csv', [{'Item': 'Máquina 1 - Setor A'}, {'Item': 'Máquina 2 - Setor B'}])
motivos_parada = load_data('motivos_parada.csv', [{'Item': 'Manutenção preventiva'}, {'Item': 'Falha técnica'}])
manutencao_feita_com = load_data('manutencao_feita_com.csv', [{'Item': 'Ferramenta 1'}, {'Item': 'Ferramenta 2'}])
emails = load_emails()

# Layout da aplicação
with st.sidebar:
    menu = option_menu(
        'Controle de Manutenção de Máquinas', 
        ['Abertura de OS', 'Fechar OS','Dashboard', 'Cadastrar Listas', 'Visualizar OS', 'Relatórios', 'Configurações', 'Histórico', 'Ajuda'],
        icons=['plus-circle', 'check-circle', 'grid', 'list', 'eye', 'file-earmark-bar-graph', 'gear', 'clock-history', 'question-circle'],
        menu_icon="cast",
        default_index=0,
        styles={
            "container": {"padding": "0!important", "background-color": "#f0f2f6"},
            "icon": {"color": "orange", "font-size": "25px"},
            "nav-link": {"font-size": "16px", "text-align": "left", "margin": "0px", "--hover-color": "#eee"},
            "nav-link-selected": {"background-color": "#ff4b4b"},
        }
    )

if menu == 'Abertura de OS':
    st.title('Abertura de Ordem de Serviço (OS)')
    numero_os = generate_os_number()
    equipamento_setor = st.selectbox('Equipamento - Setor', sorted(equipamentos_setores['Item'].unique()))
    motivo_parada = st.selectbox('Motivo da Parada', motivos_parada['Item'])
    imagem = st.file_uploader('Adicione uma Imagem', type=['jpg', 'jpeg', 'png'])
    data_input = st.date_input('Data', value=datetime.now(br_timezone).date())
    hora_input = st.time_input('Hora', value=datetime.now(br_timezone).time())
    data_hora = f"{data_input.strftime('%d/%m/%Y')} {hora_input.strftime('%H:%M')}"
    
    if st.button('Salvar'):
        image_path = None
        if imagem:
            uploads_folder = 'uploads'
            if not os.path.exists(uploads_folder):
                os.makedirs(uploads_folder)
            image_path = os.path.join(uploads_folder, imagem.name)
            with open(image_path, "wb") as f:
                f.write(imagem.getbuffer())
        
        df_os = save_os(df_os, {
            'Numero_OS': numero_os,
            'Equipamento_Setor': equipamento_setor,
            'Motivo_Parada': motivo_parada,
            'Imagem': image_path,
            'Status': 'Aberta',
            'Data_Hora': data_hora,
            'Item_Usado': '',
            'Manutencao_Com': '',
            'Data_Hora_Fechamento': ''
        }, image_path, emails)

        st.success(f'OS {numero_os} salva com sucesso!')

elif menu == 'Fechar OS':
    st.title('Fechar Ordem de Serviço (OS)')
    os_list = df_os[df_os['Status'] == 'Aberta']['Numero_OS'].tolist()
    numero_os = st.selectbox('Selecione a OS', os_list)
    item_usado = st.selectbox('Item Usado', sorted(manutencao_feita_com['Item'].unique()))
    manutencao_com = st.selectbox('Manutenção Feita com', sorted(manutencao_feita_com['Item'].unique()))
    data_input = st.date_input('Data de Fechamento', value=datetime.now(br_timezone).date())
    hora_input = st.time_input('Hora de Fechamento', value=datetime.now(br_timezone).time())
    data_hora_fechamento = f"{data_input.strftime('%d/%m/%Y')} {hora_input.strftime('%H:%M')}"

    if st.button('Salvar'):
        df_os = close_os(df_os, numero_os, item_usado, manutencao_com, data_hora_fechamento, emails)
        st.success(f'OS {numero_os} fechada com sucesso!')

elif menu == 'Cadastrar Listas':
    st.title('Cadastrar Equipamentos, Setores e Motivos de Parada')
    tab = st.selectbox('Selecionar Tabela', ['Equipamentos - Setores', 'Motivos de Parada', 'Manutenção Feita com'])
    
    if tab == 'Equipamentos - Setores':
        new_item = st.text_input('Adicionar Equipamento - Setor')
        if st.button('Adicionar'):
            equipamentos_setores = equipamentos_setores.append({'Item': new_item}, ignore_index=True)
            save_data(equipamentos_setores, 'equipamentos_setores.csv')
            st.success(f'{new_item} adicionado com sucesso!')
        st.dataframe(equipamentos_setores)
    
    elif tab == 'Motivos de Parada':
        new_item = st.text_input('Adicionar Motivo de Parada')
        if st.button('Adicionar'):
            motivos_parada = motivos_parada.append({'Item': new_item}, ignore_index=True)
            save_data(motivos_parada, 'motivos_parada.csv')
            st.success(f'{new_item} adicionado com sucesso!')
        st.dataframe(motivos_parada)
    
    elif tab == 'Manutenção Feita com':
        new_item = st.text_input('Adicionar Manutenção Feita com')
        if st.button('Adicionar'):
            manutencao_feita_com = manutencao_feita_com.append({'Item': new_item}, ignore_index=True)
            save_data(manutencao_feita_com, 'manutencao_feita_com.csv')
            st.success(f'{new_item} adicionado com sucesso!')
        st.dataframe(manutencao_feita_com)

elif menu == 'Dashboard':
    st.title('Dashboard')
    st.write('Análise gráfica das OS')
    total_os = df_os.shape[0]
    os_abertas = df_os[df_os['Status'] == 'Aberta'].shape[0]
    os_fechadas = df_os[df_os['Status'] == 'Fechada'].shape[0]
    st.write(f'Total de OS: {total_os}')
    st.write(f'OS Abertas: {os_abertas}')
    st.write(f'OS Fechadas: {os_fechadas}')
    
    df_os['Data'] = pd.to_datetime(df_os['Data_Hora'], format='%d/%m/%Y %H:%M')
    df_os['Ano'] = df_os['Data'].dt.year
    df_os['Mês'] = df_os['Data'].dt.month
    df_os['Dia'] = df_os['Data'].dt.day
    
    monthly_data = df_os.groupby(['Ano', 'Mês', 'Status']).size().unstack(fill_value=0)
    st.bar_chart(monthly_data)

elif menu == 'Visualizar OS':
    st.title('Visualizar Ordem de Serviço (OS)')
    st.dataframe(df_os)

elif menu == 'Relatórios':
    st.title('Gerar Relatórios')
    if st.button('Gerar Relatório Geral'):
        df_os.to_csv('relatorio_geral.csv', index=False)
        st.success('Relatório Geral gerado com sucesso!')
        st.write('Clique no link abaixo para baixar o relatório:')
        st.markdown(f'<a href="relatorio_geral.csv" download>Baixar Relatório Geral</a>', unsafe_allow_html=True)

elif menu == 'Configurações':
    st.title('Configurações')
    st.write('Configurar endereços de e-mail para envio de notificações:')
    new_email = st.text_input('Adicionar E-mail')
    if st.button('Adicionar'):
        emails.append(new_email)
        save_emails(emails)
        st.success(f'E-mail {new_email} adicionado com sucesso!')
    st.write('Lista de e-mails cadastrados:')
    st.write(emails)

elif menu == 'Histórico':
    st.title('Histórico')
    st.write('Histórico de OS')
    st.dataframe(df_os)

elif menu == 'Ajuda':
    st.title('Ajuda')
    st.write('Se precisar de ajuda, entre em contato com o suporte técnico.')
