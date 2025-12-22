import imaplib
import email
from email.header import decode_header
import re
from datetime import datetime
import pandas as pd
from bs4 import BeautifulSoup
import html

class ZimbraCEPExtractor:
    """Extrator de CEPs de emails do servidor Zimbra via IMAP.
    
    Esta classe permite conectar a um servidor IMAP do Zimbra, buscar emails
    em pastas específicas e extrair CEPs (Códigos de Endereçamento Postal)
    brasileiros do conteúdo das mensagens.
    
    Attributes:
        email (str): Endereço de email completo para autenticação.
        senha (str): Senha do email.
        imap_server (str): Endereço do servidor IMAP.
        mail (imaplib.IMAP4_SSL): Objeto de conexão IMAP (None se desconectado).
    
    Example:
        >>> extrator = ZimbraCEPExtractor("usuario@empresa.com.br", "senha123")
        >>> if extrator.conectar():
        ...     resultados = extrator.buscar_emails_com_cep(limite=100)
        ...     extrator.exportar_resultados(resultados)
        ...     extrator.desconectar()
    """
    
    def __init__(self, email, senha, imap_server="imap.emailzimbraonline.com"):
        """Inicializa o extrator de CEPs com credenciais de acesso.

        Args:
            email (str): Endereço de email completo (ex: usuario@empresa.com.br).
            senha (str): Senha do email para autenticação IMAP.
            imap_server (str, optional): Servidor IMAP do Zimbra. 
                Padrão: "imap.emailzimbraonline.com".
        """
        self.email = email
        self.senha = senha
        self.imap_server = imap_server
        self.mail = None

    def conectar(self):
        """Estabelece conexão SSL com o servidor IMAP do Zimbra.
        
        Realiza login usando as credenciais fornecidas na inicialização.
        A conexão é feita na porta 993 (IMAP SSL padrão).
        
        Returns:
            bool: True se a conexão foi estabelecida com sucesso, False caso contrário.
            
        Example:
            >>> extrator = ZimbraCEPExtractor("email@empresa.com", "senha")
            >>> if extrator.conectar():
            ...     print("Conectado!")
        """
        try:
            print(f"Conectando ao servidor IMAP: {self.imap_server}...")
            self.mail = imaplib.IMAP4_SSL(self.imap_server, 993)
            self.mail.login(self.email, self.senha)
            print("Conectado com sucesso!")
            return True
        except Exception as e:
            print(f"Erro ao conectar: {e}")
            return False

    def extrair_ceps(self, texto):
        """Extrai todos os CEPs válidos de um texto usando múltiplos padrões.
        
        Identifica CEPs nos seguintes formatos:
        - CEP: 12345678 (explícito sem hífen)
        - CEP:12345-678 (explícito com hífen)
        - CEP: 12345-678 (explícito com hífen e espaço)
        - 12345-678 (formato clássico)
        - 12345678 (8 dígitos quando "CEP" está próximo no texto)
        
        Args:
            texto (str): Texto do qual extrair os CEPs (pode ser None).
            
        Returns:
            list[str]: Lista de CEPs únicos encontrados no formato "XXXXX-XXX".
                Retorna lista vazia se nenhum CEP for encontrado.
                
        Note:
            - Remove duplicatas mantendo a ordem de aparição.
            - Aplica limite de 20 CEPs por texto para evitar falsos positivos.
            - Valida contexto para números de 8 dígitos (deve ter "CEP" próximo).
            
        Example:
            >>> extrator = ZimbraCEPExtractor("email@empresa.com", "senha")
            >>> ceps = extrator.extrair_ceps("CEP: 12345-678 e 98765432")
            >>> print(ceps)
            ['12345-678', '98765-432']
        """
        if not texto:
            return []

        ceps_encontrados = []
        
        # Converte para maiúsculo para facilitar busca
        texto_upper = texto.upper()
        
        # PADRÃO 1: CEP seguido de números (com ou sem hífen, com ou sem espaço)
        # Exemplos: CEP:12345678, CEP: 12345-678, CEP 12345678
        padrao_cep_explícito = r'CEP\s*:?\s*(\d{5})\s*-?\s*(\d{3})'
        matches = re.finditer(padrao_cep_explícito, texto_upper, re.IGNORECASE)
        for match in matches:
            cep = f"{match.group(1)}-{match.group(2)}"
            ceps_encontrados.append(cep)
        
        # PADRÃO 2: Números no formato XXXXX-XXX (clássico)
        padrao_com_hifen = r'\b(\d{5})\s*-\s*(\d{3})\b'
        matches = re.finditer(padrao_com_hifen, texto)
        for match in matches:
            cep = f"{match.group(1)}-{match.group(2)}"
            ceps_encontrados.append(cep)
        
        # PADRÃO 3: Sequência de exatamente 8 dígitos que PODE ser CEP
        # Mas só pega se estiver isolada (word boundary) para evitar telefones, CNPJs, etc
        padrao_8_digitos = r'\b(\d{8})\b'
        matches = re.finditer(padrao_8_digitos, texto)
        for match in matches:
            numero = match.group(1)
            # Verifica se tem "CEP" próximo (dentro de 50 caracteres antes ou depois)
            pos = match.start()
            contexto_antes = texto_upper[max(0, pos-50):pos]
            contexto_depois = texto_upper[pos:min(len(texto_upper), pos+50)]
            
            # Se tem "CEP" perto OU o número começa com dígitos típicos de CEP (01-99)
            if ('CEP' in contexto_antes or 'CEP' in contexto_depois):
                cep = f"{numero[:5]}-{numero[5:]}"
                ceps_encontrados.append(cep)
            elif numero[:2] in ['01', '02', '03', '04', '05', '06', '07', '08', '09'] or \
                (10 <= int(numero[:2]) <= 99):
                # Números que começam com 01-99 podem ser CEPs válidos
                # Mas vamos ser mais cautelosos aqui e só adicionar se não for óbvio que é outra coisa
                # Verifica se não parece data (não começa com 19, 20, 21)
                if numero[:2] not in ['19', '20', '21']:
                    cep = f"{numero[:5]}-{numero[5:]}"
                    # Adiciona apenas se ainda não temos muitos CEPs deste email
                    # (evita pegar muitos falsos positivos)
                    if len(ceps_encontrados) < 20:  # limite de segurança
                        ceps_encontrados.append(cep)
        
        # Remove duplicatas mantendo ordem
        ceps_unicos = list(dict.fromkeys(ceps_encontrados))
        
        return ceps_unicos

    def limpar_html(self, html_text):
        """Remove tags HTML e extrai apenas o texto visível.
        
        Processa conteúdo HTML removendo scripts, estilos e tags, retornando
        apenas o texto limpo e legível.
        
        Args:
            html_text (str): String contendo código HTML (pode ser None).
            
        Returns:
            str: Texto limpo sem tags HTML, com espaços normalizados.
                Retorna string vazia se input for None ou vazio.
                
        Note:
            - Remove completamente tags <script> e <style>.
            - Decodifica entidades HTML (ex: &nbsp; vira espaço).
            - Normaliza múltiplos espaços em um único espaço.
            
        Example:
            >>> extrator = ZimbraCEPExtractor("email@empresa.com", "senha")
            >>> html = "<p>CEP: <b>12345-678</b></p>"
            >>> texto = extrator.limpar_html(html)
            >>> print(texto)
            'CEP: 12345-678'
        """
        if not html_text:
            return ""

        # Cria o objeto BeautifulSoup
        soup = BeautifulSoup(html_text, "html.parser")

        # Remove scripts e styles
        for tag in soup(["script", "style"]):
            tag.decompose()

        # Pega só o texto visível
        texto = soup.get_text(separator=" ")

        # Decodifica entidades HTML
        texto = html.unescape(texto)

        # Normaliza espaços
        texto = " ".join(texto.split())

        return texto.strip()

    def decodificar_conteudo(self, msg):
        """Decodifica e extrai o conteúdo completo de uma mensagem de email.
        
        Processa emails multipart e simples, priorizando texto simples sobre HTML.
        Ignora anexos e decodifica diferentes charsets automaticamente.
        
        Args:
            msg (email.message.Message): Objeto de mensagem de email do módulo email.
            
        Returns:
            str: Conteúdo textual completo do email. Se houver partes text/plain,
                retorna essas. Caso contrário, retorna HTML limpo. String vazia
                se não houver conteúdo.
                
        Note:
            - Prioriza text/plain sobre text/html.
            - Tenta múltiplos charsets: específico da parte, utf-8, latin-1.
            - Anexos são automaticamente ignorados.
            - HTML é convertido para texto limpo via limpar_html().
            
        Example:
            >>> import email
            >>> msg = email.message_from_bytes(raw_email_data)
            >>> conteudo = extrator.decodificar_conteudo(msg)
        """
        # Listas para armazenar texto simples e HTML separadamente
        conteudo_texto = []
        conteudo_html = []

        # Verifica se o email tem múltiplas partes (texto + HTML + anexos)
        if msg.is_multipart():
            # Percorre todas as partes do email
            for parte in msg.walk():
                ctype = parte.get_content_type()  # Tipo MIME da parte (ex: text/plain)
                disp = str(parte.get("Content-Disposition") or "")  # Se é anexo ou inline
                
                # Ignora anexos
                if "attachment" in disp.lower():
                    continue

                # Tenta pegar o conteúdo decodificado (Base64 ou Quoted-Printable)
                try:
                    payload = parte.get_payload(decode=True)
                except Exception:
                    payload = None

                # Pula se não houver conteúdo
                if not payload:
                    continue

                # Define charset da parte (padrão 'utf-8')
                charset = parte.get_content_charset() or 'utf-8'
                try:
                    texto = payload.decode(charset, errors="ignore")  # Converte bytes para string
                except Exception:
                    # Se falhar, tenta utf-8 e latin-1 como fallback
                    try:
                        texto = payload.decode('utf-8', errors="ignore")
                    except Exception:
                        try:
                            texto = payload.decode('latin-1', errors="ignore")
                        except Exception:
                            texto = ""  # Último recurso: vazio

                # Guarda em listas separadas por tipo
                if ctype == "text/plain":
                    conteudo_texto.append(texto)
                elif ctype == "text/html":
                    conteudo_html.append(texto)

        else:
            # Se não for multipart, trata a mensagem como única parte
            try:
                payload = msg.get_payload(decode=True)
                if payload:
                    charset = msg.get_content_charset() or 'utf-8'
                    try:
                        texto = payload.decode(charset, errors="ignore")
                    except Exception:
                        try:
                            texto = payload.decode('utf-8', errors="ignore")
                        except Exception:
                            texto = payload.decode('latin-1', errors="ignore")
                    
                    ctype = msg.get_content_type()  # Tipo MIME da mensagem
                    # Armazena no lugar certo
                    if ctype == "text/plain":
                        conteudo_texto.append(texto)
                    elif ctype == "text/html":
                        conteudo_html.append(texto)
                    else:
                        conteudo_texto.append(texto)  # Caso não seja texto ou HTML
            except Exception:
                pass  # Ignora falhas silenciosamente

        # Retorna texto plano se disponível
        if conteudo_texto:
            return "\n".join(conteudo_texto).strip()
        
        # Senão, retorna HTML limpo (chama método que você implementou)
        if conteudo_html:
            html_total = "\n".join(conteudo_html)
            return self.limpar_html(html_total)
        
        # Se não tiver nada, retorna string vazia
        return ""

    def buscar_emails_com_cep(self, limite=None, data_inicio=None, pastas=None):
        """Busca e processa emails que contêm CEPs em pastas específicas.
        
        Varre emails em busca de CEPs no assunto ou corpo, filtrando apenas
        mensagens que mencionam "CEP" ou "CEPS". Suporta filtro por data e
        limite de emails processados.

        Args:
            limite (int, optional): Número máximo de emails a processar por pasta.
                None = sem limite. Padrão: None.
            data_inicio (str, optional): Data inicial para filtro no formato IMAP
                "DD-Mon-YYYY" (ex: "01-Jan-2025"). None = sem filtro de data.
                Padrão: None.
            pastas (list[str], optional): Lista de nomes de pastas para buscar
                (ex: ["INBOX", "Area Restrita"]). None = busca em todas as pastas.
                Padrão: None.

        Returns:
            list[dict]: Lista de dicionários, cada um contendo:
                - data (str): Data do email.
                - remetente (str): Endereço do remetente.
                - assunto (str): Assunto truncado (max 100 caracteres).
                - corpo (str): Conteúdo completo do email.
                - ceps (str): CEPs encontrados separados por vírgula.
                - quantidade (int): Número de CEPs únicos encontrados.
                
        Note:
            - Processa apenas emails que mencionam "CEP" ou "CEPS".
            - Se pastas=None, busca em todas as pastas da conta.
            - Imprime progresso e estatísticas durante a execução.
            - Erros individuais são logados mas não interrompem a busca.
            
        Example:
            >>> extrator = ZimbraCEPExtractor("email@empresa.com", "senha")
            >>> extrator.conectar()
            >>> resultados = extrator.buscar_emails_com_cep(
            ...     limite=100,
            ...     data_inicio="01-Jan-2025",
            ...     pastas=["INBOX"]
            ... )
            >>> print(f"Encontrados {len(resultados)} emails com CEP")
        """
        resultados = []
        try:
            # Lista todas as pastas se não foram passadas
            if pastas is None:
                status, todas_pastas = self.mail.list()
                if status != "OK" or not todas_pastas:
                    print("Erro ao listar pastas ou nenhuma pasta encontrada.")
                    return resultados

                pastas = []
                for p in todas_pastas:
                    partes = p.decode().split(' "/" ')
                    if len(partes) == 2:
                        pastas.append(partes[1].strip('"'))

            print(f"📁 Pastas que serão processadas: {len(pastas)}")

            emails_com_cep = 0
            emails_sem_cep = 0

            for pasta_nome in pastas:
                try:
                    status_select, _ = self.mail.select(f'"{pasta_nome}"')
                    if status_select != "OK":
                        print(f"Erro ao selecionar pasta '{pasta_nome}': {status_select}")
                        continue
                except Exception as e:
                    print(f"✗ Não foi possível selecionar a pasta '{pasta_nome}': {e}")
                    continue

                # Monta filtro IMAP
                if data_inicio:
                    filtro_imap = f'SINCE "{data_inicio}"'
                else:
                    filtro_imap = 'ALL'

                status, messages = self.mail.search(None, filtro_imap)
                if status != "OK" or not messages or not messages[0]:
                    continue

                email_ids = messages[0].split()
                total = len(email_ids)

                # Aplica limite por pasta
                if limite and isinstance(limite, int) and limite > 0:
                    email_ids = email_ids[-limite:]

                for i, email_id in enumerate(email_ids, 1):
                    try:
                        status, msg_data = self.mail.fetch(email_id, "(RFC822)")
                        if status != "OK" or not msg_data or not msg_data[0]:
                            continue

                        msg = email.message_from_bytes(msg_data[0][1])

                        # Decodifica assunto
                        raw_subject = msg.get("Subject", "")
                        decoded_fragments = decode_header(raw_subject)
                        subject_parts = []
                        for frag, enc in decoded_fragments:
                            if isinstance(frag, bytes):
                                try:
                                    subject_parts.append(frag.decode(enc or 'utf-8', errors="ignore"))
                                except Exception:
                                    try:
                                        subject_parts.append(frag.decode('utf-8', errors="ignore"))
                                    except Exception:
                                        subject_parts.append(frag.decode('latin-1', errors="ignore"))
                            else:
                                subject_parts.append(str(frag))
                        subject = " ".join(part for part in subject_parts if part).strip()

                        from_ = msg.get("From", "")
                        date = msg.get("Date", "")

                        # Extrai conteúdo
                        conteudo = self.decodificar_conteudo(msg)

                        # Só processa CEPs se assunto ou corpo mencionar CEP/CEPS
                        if re.search(r'\bCEPS?\b', subject, re.IGNORECASE) or re.search(r'\bCEP\b', conteudo, re.IGNORECASE):
                            texto_completo = f"{subject}\n{conteudo}"
                            ceps = self.extrair_ceps(texto_completo)
                        else:
                            ceps = []

                        if ceps:
                            emails_com_cep += 1
                            resultados.append({
                                'data': date,
                                'remetente': from_,
                                'assunto': subject[:100],
                                'corpo': conteudo,
                                'ceps': ', '.join(ceps),
                                'quantidade': len(ceps)
                            })
                            print(f"[{emails_com_cep}/{emails_com_cep + emails_sem_cep}] ✓ {len(ceps)} CEP(s): {', '.join(ceps[:3])}{'...' if len(ceps) > 3 else ''}")
                        else:
                            emails_sem_cep += 1
                            if (emails_com_cep + emails_sem_cep) % 10 == 0:
                                print(f"[{emails_com_cep + emails_sem_cep}] - Processando... ({emails_com_cep} com CEP até agora)")

                    except Exception as e:
                        print(f"✗ Erro ao processar email: {e}")
                        continue

            print(f"\n📊 Estatísticas finais:")
            print(f"   • Emails com CEP: {emails_com_cep}")
            print(f"   • Emails sem CEP: {emails_sem_cep}")

            return resultados

        except Exception as e:
            print(f"Erro ao buscar emails: {e}")
            return resultados

        except Exception as e:
            print(f"Erro ao buscar emails: {e}")
            return resultados

    def desconectar(self):
        """Encerra a conexão com o servidor IMAP de forma segura.
        
        Realiza logout do servidor IMAP e libera os recursos de conexão.
        Erros durante a desconexão são capturados e logados.
        
        Note:
            - Pode ser chamado mesmo se a conexão já estiver fechada.
            - Recomendado usar sempre após terminar o processamento.
            
        Example:
            >>> extrator = ZimbraCEPExtractor("email@empresa.com", "senha")
            >>> extrator.conectar()
            >>> # ... processar emails ...
            >>> extrator.desconectar()
        """
        try:
            if self.mail:
                self.mail.logout()
                print("Desconectado do servidor IMAP.")
        except Exception as e:
            print("Erro ao desconectar:", e)

    def exportar_resultados(self, resultados, formato='excel'):
        """Exporta os resultados da busca para arquivo Excel ou CSV.
        
        Cria um arquivo com timestamp contendo todos os CEPs encontrados
        organizados em tabela com informações do email.

        Args:
            resultados (list[dict]): Lista de dicionários retornada por
                buscar_emails_com_cep().
            formato (str, optional): Formato do arquivo de saída.
                Valores aceitos: 'excel' ou 'csv'. Padrão: 'excel'.

        Returns:
            str: Nome do arquivo criado (ex: "ceps_extraidos_20250101_123045.xlsx").
                None se não houver resultados para exportar.
                
        Note:
            - Arquivo é salvo no diretório atual.
            - Nome inclui timestamp para evitar sobrescrita.
            - CSV usa encoding UTF-8 com BOM para compatibilidade com Excel.
            - Colunas: data, remetente, assunto, corpo, ceps, quantidade.
            
        Example:
            >>> resultados = extrator.buscar_emails_com_cep(limite=100)
            >>> arquivo = extrator.exportar_resultados(resultados, formato='excel')
            >>> print(f"Dados salvos em: {arquivo}")
        """
        if not resultados:
            print("Nenhum resultado para exportar.")
            return None
        
        df = pd.DataFrame(resultados)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        if formato == 'excel':
            filename = f"ceps_extraidos_{timestamp}.xlsx"
            df.to_excel(filename, index=False)
            print(f"\n✓ Resultados exportados para: {filename}")
        else:
            filename = f"ceps_extraidos_{timestamp}.csv"
            df.to_csv(filename, index=False, encoding='utf-8-sig')
            print(f"\n✓ Resultados exportados para: {filename}")

        return filename