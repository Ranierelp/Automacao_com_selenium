import os
import time
from datetime import date
from imap_tools import MailBox, AND
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from decouple.decouple import config
from reportlab.lib.pagesizes import A4
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle, Paragraph, Frame, PageTemplate, NextPageTemplate)
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

buscar_cursos = (By.CSS_SELECTOR, "input[placeholder='Buscar cursos']")
button_pesquiser = (By.CSS_SELECTOR, "button[type='submit']")
config_disciplina = (By.CSS_SELECTOR, "a[class='action-edit']")
button_geral = (By.CSS_SELECTOR, "a[data-action='togglecourseindexsection']")
ativar_modo_edicao = (By.CSS_SELECTOR, "input[name= 'setmode']")
adicionar_atividade = (By.CSS_SELECTOR, "button[data-action='open-chooser']")
ferramenta_externa = (By.CSS_SELECTOR, "a[title='Adicionar um novo Ferramenta externa']")
colocar_nome_da_ua = (By.XPATH, "//*[@id='id_name']")
mostrar_mais = (By.XPATH, "/html/body/div[6]/div[6]/div/div[2]/div/section/div/form/fieldset[1]/div[2]/div[2]/div/a")
colocar_link_ua = (By.XPATH, "//*[@id='id_toolurl']")
salvar_continuar_curso = (By.CSS_SELECTOR, "input[value='Salvar e voltar ao curso']")
titulo_da_disciplina = (By.CSS_SELECTOR, "h1[class='h2']")

campo_login = (By.ID, "username")
campo_senha = (By.ID, "password")
login_button = (By.ID, "loginbtn")

diretorio_atual = os.path.dirname(__file__)
nome_arquivo_pdf = os.path.join(diretorio_atual, "relatorio.pdf")
caminho_imagem_papel_timbrado = os.path.join(diretorio_atual, "stylefolha", "Timbrado UNIFIP 2.png")

def fazerLogin(usuario, senha):   
        
    escrever(campo_login, usuario)
    escrever(campo_senha, senha)
    clicar(login_button)

#---------------------------------------------------------------------------------#

def encontrar_elemento(locator):
    return driver.find_element(*locator)
    
def encontrar_elementos(locator):
    return driver.find_elements(*locator)

def escrever(locator, text):
    encontrar_elemento(locator).send_keys(text)

def clicar( locator):
    encontrar_elemento(locator).click()

def verificar_se_elemento_existe(locator):
    assert encontrar_elemento(locator).is_displayed(), f"o elemento {locator} não foi encontrado"

def pegar_elemento_texto(locator):
    return encontrar_elemento(locator).text

def esperar_elemento_aparecer(locator, timeOut=5):
    return WebDriverWait(driver, timeOut).until(EC.presence_of_element_located(locator))
    
def rolar_para_elemento(locator):
    elemento = encontrar_elemento(locator)
    driver.execute_script("arguments[0].scrollIntoView();", elemento)
    
def encontrar_select(by, value):
    return driver.find_element(by, value)

def verificacao_de_um_elemento(locator,texto):
    encontrado = encontrar_elemento(locator).text
    assert encontrado == texto, f"o elemento encontrado foi: ({encontrado}) e o elemento esperado foi: ({texto})"

#---------------------------------------------------------------------------------#

def setup():
    global driver
    driver = webdriver.Chrome()
    driver.implicitly_wait(5)
    driver.maximize_window()
    driver.get("https://presencial.unifipdigital.com.br/login/index.php")
    
def teardown():
    driver.quit()

#---------------------------------------------------------------------------------#

def criar_PDF(dados_relatorio, nome_arquivo, quantidade_uas):

    def myFirstPage(canvas, doc):
        imagem_papel_timbrado = caminho_imagem_papel_timbrado
        width, height = A4
        canvas.drawImage(imagem_papel_timbrado, 0, 0, width=width, height=height)

    def myLaterPages(canvas, doc):
        imagem_papel_timbrado = caminho_imagem_papel_timbrado
        width, height = A4
        canvas.drawImage(imagem_papel_timbrado, 0, 0, width=width, height=height)

    doc = SimpleDocTemplate(nome_arquivo, pagesize=A4)
    frame = Frame(doc.leftMargin - 30, doc.bottomMargin + 15, doc.width +60, doc.height -85, showBoundary=0)
    watermark_template = PageTemplate(id='Later' ,frames=[frame], onPage=myLaterPages)
    conteudo_PDF = []  
    styles = getSampleStyleSheet()

    title = "Relatório de Disciplinas Cadastradas"
    conteudo_PDF.append(Paragraph(title, getSampleStyleSheet()["Title"]))
    
    table_data = [["Disciplina", "Status"]]  
    for disciplina, status in dados_relatorio.items():
        text_color = colors.black if status == "Cadastrada com sucesso" else colors.red
        
        row = [Paragraph(f"<font color={text_color}>{disciplina}</font>", styles['Normal']),
              Paragraph(f"<font color={text_color}>{status}</font>", styles['Normal'])]
        table_data.append(row)

    table = Table(table_data)
    table.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#003363')),
                               ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                               ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                               ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                               ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                               ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                               ('GRID', (0, 0), (-1, -1), 1, colors.black)]))
    
    conteudo_PDF.append(table)
    conteudo_PDF.append(NextPageTemplate('Later'))
    doc.addPageTemplates(watermark_template)
    total_disciplinas = len(dados_relatorio)
    total_paragraph = Paragraph(f"Total de Disciplinas: {total_disciplinas}<br/>Total de UAs Cadastradas: {quantidade_uas}", styles['Normal'])
    conteudo_PDF.append(total_paragraph)
    doc.build(conteudo_PDF, onFirstPage=myFirstPage, onLaterPages=myLaterPages)
#---------------------------------------------------------------------------------#

def automocaao_email():

    setup() 

    relatorio = {}
    user = config('USER')
    password = config('PASSWORD')

    email_ = config('EMAIL')
    senha = config('SENHA')

    fazerLogin(user, password)

    meu_email = MailBox("imap.gmail.com").login(email_, senha)
    data_inicial = date(2023,8,11)
    lista_email = meu_email.fetch(AND(from_ = "no-reply@grupoa.education", text="23.2 aprovada ", date_gte = data_inicial,seen=False))

    quantidade_disciplina = 0
    quantidade_uas = 0
    
    clicar(ativar_modo_edicao)

    for email in lista_email:
        dados_email = email.subject

        inicio_disciplina = "Disciplina "
        fim_disciplina = " aprovada."
        nome_disciplina = dados_email[dados_email.find(inicio_disciplina) + len(inicio_disciplina):dados_email.find(fim_disciplina)]

        time.sleep(2)
        driver.get("https://presencial.unifipdigital.com.br/course/management.php")
        escrever(buscar_cursos, nome_disciplina)
        clicar(button_pesquiser)
        time.sleep(5)
        
        try:  
            esperar_elemento_aparecer(config_disciplina)
            clicar(config_disciplina)
            verificacao_de_um_elemento(titulo_da_disciplina,nome_disciplina)
            clicar(button_geral)
            clicar(adicionar_atividade)
            clicar(ferramenta_externa)

            relatorio[nome_disciplina] = "Cadastrada com sucesso"

        except:
            relatorio[nome_disciplina] = "Não encontrada"

        if relatorio[nome_disciplina] == "Cadastrada com sucesso":
            if email.html:
                soup = BeautifulSoup(email.html,  "html.parser")
                texto_html = soup.get_text()
                dados_corpo_email = texto_html

                inicio_ua = "UA's:"
                fim_ua = "SAGAH"
                info_ua = dados_corpo_email[dados_corpo_email.find(inicio_ua) + len(inicio_ua): dados_corpo_email.find(fim_ua)]
                uas = info_ua.strip().split('\n\n')
                contador = len(uas)
                
                for ua in uas:
                    nome_ua = ua.strip().split('\n')[0]
                    link_ua = ua.strip().split('\n')[1]
                    
                    escrever(colocar_nome_da_ua, nome_ua)
                    clicar(mostrar_mais)
                    
                    select_modo_de_vizu = encontrar_select(By.XPATH, "//*[@id='id_launchcontainer']")
                    select = Select(select_modo_de_vizu)
                    select.select_by_visible_text("Nova janela")
                
                    escrever(colocar_link_ua, link_ua)
                    
                    quantidade_uas += 1
                    clicar(salvar_continuar_curso)
                
                    contador -= 1
                    if contador == 0:
                        break
                        
                    else:
                        clicar(adicionar_atividade)
                        clicar(ferramenta_externa)
                

        quantidade_disciplina += 1
    teardown()
    criar_PDF(relatorio,nome_arquivo_pdf, quantidade_uas)

automocaao_email()
