import os
import shutil
from tkinter import *
from tkinter import messagebox, filedialog
from datetime import datetime
import logging
from typing import Tuple, Optional, List

import pandas as pd

class DiretorioManager:
    def __init__(self):
        self._diretorio_base: Optional[str] = None
        self._setup_logging()

    def _setup_logging(self):
        logging.basicConfig(
            filename='gerenciador_pastas.log',
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )

    @property
    def diretorio_base(self) -> Optional[str]:
        return self._diretorio_base

    @diretorio_base.setter
    def diretorio_base(self, valor: str):
        if not os.path.isdir(valor):
            raise ValueError("Diretório inválido")
        if not os.access(valor, os.W_OK):
            raise PermissionError("Sem permissão de escrita no diretório")
        self._diretorio_base = valor
        logging.info(f"Diretório base configurado: {valor}")

    def buscar_pastas_equipamento(self, nome_equipamento: str,tensao:str) -> List[Tuple[str, float]]:
        """Busca todas as pastas que contêm o nome do equipamento."""
        if not self._diretorio_base:
            raise ValueError("Diretório base não configurado")
        if not nome_equipamento.strip():
            raise ValueError("Nome do equipamento não pode estar vazio")
            
        pastas_encontradas = []

        for pasta in os.listdir(self._diretorio_base):
            if nome_equipamento.lower() in pasta.lower() and tensao.lower() in pasta.lower():  # Simplificado para uma única verificação
                caminho_completo = os.path.join(self._diretorio_base, pasta)
                if os.path.isdir(caminho_completo):
                    data_modificacao = os.path.getmtime(caminho_completo)
                    pastas_encontradas.append((pasta, data_modificacao))
        return pastas_encontradas

class EquipamentoManager:

    def __init__(self, diretorio_manager: DiretorioManager):
        self._diretorio_manager = diretorio_manager

    def validar_dados(self, nome_equipamento: str, numero_os: str, tensao: str) -> None:
        """Valida os dados de entrada."""
        if not all([nome_equipamento.strip(), numero_os.strip(), tensao.strip()]):
            raise ValueError("Todos os campos são obrigatórios")
        
        if not numero_os.isdigit():
            raise ValueError("Número da OS deve conter apenas dígitos")

        try:
            float(tensao.replace('V', '').strip())
        except ValueError:
            raise ValueError("Tensão deve ser um número válido")


       #Retorna a OS,o nome do equipamento e a tensão
    def _buscaInformacoesEquipamentoOS(self, arquivo: str) -> Optional [tuple[str,str,str]]:

        df = pd.read_excel(arquivo)
        informacoes = []
        try:
            #Percorre as colunas de nome do equipamento,num da OS e tensão para criar uma lista com as informações
            for equipamento in df['Equipamento'].dropna():

                    for tensao in df['Tensão']:

                        for os in df['OS']:

                            informacoes.append((os,equipamento,tensao))


            return informacoes
        
        except Exception as e:
            logging.error(f"Erro ao buscar informações do equipamento: {str(e)}")

    def obter_pasta_mais_recente(self, nome_equipamento: str,tensao:str) -> Optional[Tuple[str, float]]:
        #????
        informacoes = self._buscaInformacoesEquipamentoOS(self._diretorio_manager.arquivo_excel)

  
        """Obtém a pasta mais recente para um dado equipamento."""
        pastas = self._diretorio_manager.buscar_pastas_equipamento(nome_equipamento,tensao)
        if not pastas:
            return None
        return max(pastas, key=lambda x: x[1])

    def criar_nova_pasta(self, nome_equipamento: str, numero_os: str, 
                        tensao: str, pasta_origem: Tuple[str, float]) -> Tuple[str, float]:
        """Cria uma nova pasta para o equipamento com base em uma pasta existente."""
        self.validar_dados(nome_equipamento, numero_os, tensao)
        
        if not pasta_origem:
            raise ValueError("Pasta de origem não encontrada")

        nova_pasta = f"OS{numero_os}-{nome_equipamento} {tensao}"
        origem = os.path.join(self._diretorio_manager.diretorio_base, pasta_origem[0])
        destino = os.path.join(self._diretorio_manager.diretorio_base, nova_pasta)

        if not os.path.exists(origem):
            raise FileNotFoundError("Pasta de origem não existe mais")
            
        if os.path.exists(destino):
            raise FileExistsError("Já existe uma pasta com este nome")
        
        try:
            shutil.copytree(origem, destino)
            return nova_pasta, os.path.getmtime(destino)
        except Exception as e:
            logging.error(f"Erro ao criar pasta: {str(e)}")
            raise

class InterfaceGrafica:
    def __init__(self):
        self._diretorio_manager = DiretorioManager()
        self._equipamento_manager = EquipamentoManager(self._diretorio_manager)
        self._criar_interface()

    def _criar_interface(self):
        self.janela = Tk()
        self.janela.title("Gerenciador de Pastas de Equipamentos")
        self.janela.geometry("500x300")
        self.janela.protocol("WM_DELETE_WINDOW", self._ao_fechar)

        self._criar_componentes()
        self.janela.mainloop()

    def _criar_componentes(self):
        # Frame principal
        main_frame = Frame(self.janela)
        main_frame.pack(padx=20, pady=20, fill=BOTH, expand=True)

        Button(main_frame, text="Selecionar Diretório", 
               command=self._selecionar_diretorio).pack(pady=10)
            
        Button(main_frame, text="Selecionar Arquivo Excel", 
               command=self._selecionarArquivoExcelOS).pack(pady=10)

        Button(main_frame, text="Crias pastas dos equipamentos", 
               command=self._selecionarArquivoExcelOS).pack(pady=10)

        self.label_diretorio = Label(main_frame, text="Nenhum diretório selecionado", 
                                   wraplength=400)
        self.label_diretorio.pack(pady=5)

        # Frame para entradas
        entry_frame = Frame(main_frame)
        entry_frame.pack(pady=10, fill=X)

        campos = [
            ("Nome do Equipamento:", "entrada_nome_equipamento"),
            ("Número da OS:", "entrada_número_os"),
            ("Tensão:", "entrada_tensão")
        ]

        for label_text, attr_name in campos:
            Label(entry_frame, text=label_text).pack(pady=5)
            entry = Entry(entry_frame, width=40)
            entry.pack()
            setattr(self, attr_name, entry)

        self.botao_criar = Button(main_frame, text="Criar Nova Pasta", 
                                command=self._processar_criacao, state='disabled')
        self.botao_criar.pack(pady=20)

    def _ao_fechar(self):
        """Método chamado quando a janela é fechada."""
        if messagebox.askokcancel("Sair", "Deseja realmente sair?"):
            self.janela.destroy()

    def _selecionar_diretorio(self):
        try:
            diretorio = filedialog.askdirectory(title="Selecione a pasta dos equipamentos")
            if diretorio:
                self._diretorio_manager.diretorio_base = diretorio
                self.label_diretorio.config(text=f"Diretório: {diretorio}")
                self.botao_criar.config(state='normal')
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    def _selecionarArquivoExcelOS(self):
        try:
            arquivo = filedialog.askopenfilename(title="Selecione o arquivo Excel com as OS",filetypes=[("Excel files", "*.xlsx")])
            if arquivo:
                self._diretorio_manager.arquivo_excel = arquivo
                self.label_arquivo_excel.config(text=f"Equipamentos Mês: {arquivo}")
                self.botao_criar.config(state='normal')
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    def _processar_criacao(self):
        try:
            nome_equip = self.entrada_nome_equipamento.get()
            numero_os = self.entrada_número_os.get()
            tensao = self.entrada_tensão.get()

            # Validação dos dados
            self._equipamento_manager.validar_dados(nome_equip, numero_os, tensao)

            pasta_recente = self._equipamento_manager.obter_pasta_mais_recente()
            # pasta_recente = self._equipamento_manager.obter_pasta_mais_recente(nome_equip,tensao)
            if not pasta_recente:
                messagebox.showwarning("Aviso", 
                                     f"Não foi encontrada pasta para: {nome_equip}")
                return

            nova_pasta, data_mod = self._equipamento_manager.criar_nova_pasta(
                nome_equip, numero_os, tensao, pasta_recente)

            self._mostrar_sucesso(nova_pasta, pasta_recente[0], data_mod)
            self._limpar_campos()

        except Exception as e:
            messagebox.showerror("Erro", str(e))

   


    def _mostrar_sucesso(self, nova_pasta: str, pasta_origem: str, data_mod: float):
        data_formatada = datetime.fromtimestamp(data_mod).strftime('%d/%m/%Y %H:%M:%S')
        messagebox.showinfo("Sucesso", 
                          f"Pasta criada: {nova_pasta}\n" + 
                          f"Origem: {pasta_origem}\n" +
                          f"Data modificação: {data_formatada}")

    def _limpar_campos(self):
        for campo in [self.entrada_nome_equipamento, self.entrada_número_os, 
                     self.entrada_tensão]:
            campo.delete(0, END)

if __name__ == "__main__":
    app = InterfaceGrafica()
