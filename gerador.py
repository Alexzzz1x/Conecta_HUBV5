import os
import qrcode
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.graphics.barcode import code128

def gerar_pdf_placa_tradicional(deposito, material, descricao, qtd, unidade, lote="", tamanho="Grande", qr_topo=""):
    # Configuração de diretórios
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))
    pasta_destino = os.path.join(diretorio_atual, "etiquetas_geradas")
    if not os.path.exists(pasta_destino): os.makedirs(pasta_destino)
    
    # 1. Definição do tamanho e escala
    if tamanho == "Pequena":
        page_size = (400, 250)
        escala = 0.45 
    elif tamanho == "Média":
        page_size = (600, 350)
        escala = 0.62
    else: # Grande
        page_size = landscape(A4)
        escala = 1.0 
        
    nome_arquivo = str(material).strip() or "identificacao_livre"
    caminho_salvamento = os.path.join(pasta_destino, f"etiqueta_{nome_arquivo}.pdf")
    c = canvas.Canvas(caminho_salvamento, pagesize=page_size)
    l, a = page_size

    # 2. Desenho da Estrutura
    c.setStrokeColor(colors.black); c.setLineWidth(2)
    c.rect(20 * escala, 20 * escala, l - (40 * escala), a - (40 * escala))
    c.line(20 * escala, a - (100 * escala), l - (20 * escala), a - (100 * escala))
    
    # 3. Logo
    caminho_logo = os.path.join(diretorio_atual, "logo_cntc.png")
    if os.path.exists(caminho_logo):
        c.drawImage(caminho_logo, 30 * escala, a - (85 * escala), width=150 * escala, height=50 * escala, preserveAspectRatio=True, anchor='sw')

    # 4. QR Code do Canto Superior Direito (Se preenchido)
    c.setFont("Helvetica-Bold", 20 * escala)
    if qr_topo:
        img_qr_topo = qrcode.make(str(qr_topo))
        caminho_qr_topo_temp = os.path.join(pasta_destino, f"temp_qr_topo_{nome_arquivo}.png")
        img_qr_topo.save(caminho_qr_topo_temp)
        
        tamanho_qr_topo = 65 * escala
        pos_x_topo = l - (30 * escala) - tamanho_qr_topo
        pos_y_topo = a - (25 * escala) - tamanho_qr_topo
        
        c.drawImage(caminho_qr_topo_temp, pos_x_topo, pos_y_topo, width=tamanho_qr_topo, height=tamanho_qr_topo)
        
        if os.path.exists(caminho_qr_topo_temp):
            os.remove(caminho_qr_topo_temp)
            
        # Desloca o texto do depósito para a esquerda do QR Code para não sobrepor
        c.drawRightString(pos_x_topo - (15 * escala), a - (65 * escala), f"DEPÓSITO: {deposito}")
    else:
        c.drawRightString(l - (30 * escala), a - (65 * escala), f"DEPÓSITO: {deposito}")
    
    # 5. Textos Principais
    c.setFont("Helvetica-Bold", 110 * escala)
    c.drawCentredString(l/2, 370 * escala, deposito)
    
    c.setFont("Helvetica-Bold", 55 * escala)
    c.drawCentredString(l/2, 290 * escala, str(material))
    
    c.setFont("Helvetica-Bold", 24 * escala)
    c.drawCentredString(l/2, 230 * escala, descricao[:70])
    
    if qtd:
        c.setFont("Helvetica-Bold", 38 * escala)
        c.drawCentredString(l/2, 160 * escala, f"QUANTIDADE: {qtd} {unidade}")
    
    if lote:
        c.setFont("Helvetica-Bold", 24 * escala)
        c.drawCentredString(l/2, 110 * escala, f"LOTE: {lote}")

    # 6. Código de Barras Tradicional (Embaixo)
    if str(material).strip():
        codigo_barras = code128.Code128(str(material), barHeight=65 * escala, barWidth=2.2 * escala)
        codigo_barras.drawOn(c, (l - codigo_barras.width) / 2, 30 * escala)

    c.showPage()
    c.save()
    return caminho_salvamento
