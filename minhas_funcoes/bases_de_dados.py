from dataclasses import dataclass, field
from google_api_functions import get_values
import pandas as pd
import datetime
import re
import numpy as np
import os

@dataclass
class baseDeDadosGoogle:
    """Classe para lidar com interações com bases de planilhas do Google"""

    baseVUC: object = None
    baseInventario: object = None
    baseDevolucao: object = None
    baseExpedicaoVUC: object = None
    dicionarioDeBases: list[str] = field(default_factory=list) 
    baseConcatenada: object = None
    contagemProgresso: int = 0

    def mapear_bases_a_manipular(self) -> None:
        """Mapea nome de bases a manipular"""

        self.dicionarioDeBases = ['baseVUC','baseInventario','baseDevolucao','baseExpedicaoVUC']

    def baixar_bases(self) -> None:
        """Baixa bases do Google Sheets de acordo com ID fornecido"""

        for i in range(5):
            try:
                self.baseInventario = get_values('1JcQH_Df-_jnFMGJoFXVk4yLr43mQTo5elD2VObd2iOw', "'Respostas ao formulário 1'!A1:D")
                self.baseVUC = get_values('1GIlraHslu0FahZK6-KLEE4GY7odiIMHfg9TKTv3z0FI', "'Respostas ao formulário 1'!A1:C")
                self.baseDevolucao = get_values('1jK_jw6vhLc03MLJyfdZ2nNqiTYal8zyDdngM_kOQ1ZM', "'Respostas ao formulário 1'!A1:C")
                self.baseExpedicaoVUC = get_values('1pPOhGrHQBm8Rac09b9yhrCULD0nA5Czo8Wvsp0Lh47k', "'Respostas ao formulário 1'!A1:E")
                break
            except: pass

    def quebrar_ids_por_linhas(self,base) -> list:
        """Quebra ids de pacotes por linhas e conserva dados adjacentes"""

        for i,j in zip(enumerate(base[0]), base[0]):
            if "Shipment" in j: 
                indice_coluna_shipments, _ = i
        
        indice_colunas_adjacentes_a_shipments = [i for i,_ in enumerate(base[0]) if i != indice_coluna_shipments]
        
        linhas = []

        for linha in base[1:]:
            while len(linha) < len(base[0]): linha.append('')
            ids = [re.search('4\d\d\d\d\d\d\d\d\d\d', i)[0] for i in linha[indice_coluna_shipments].split('\n') if str(type(re.search('4\d\d\d\d\d\d\d\d\d\d', i)))!="<class 'NoneType'>"]
            for id_ in ids:
                linha = [linha[coluna].upper() for coluna in indice_colunas_adjacentes_a_shipments]
                linha.insert(indice_coluna_shipments,id_)
                linhas.append(linha)

        return linhas
            
    def tratar_bases(self) -> None:
        """Trata bases e transforma em tabela"""

        self.baseInventario = pd.DataFrame(self.quebrar_ids_por_linhas(self.baseInventario), columns=self.baseInventario[0])
        self.baseVUC = pd.DataFrame(self.quebrar_ids_por_linhas(self.baseVUC), columns=self.baseVUC[0])
        self.baseDevolucao = pd.DataFrame(self.quebrar_ids_por_linhas(self.baseDevolucao), columns=self.baseDevolucao[0])
        self.baseExpedicaoVUC = pd.DataFrame(self.quebrar_ids_por_linhas(self.baseExpedicaoVUC), columns=self.baseExpedicaoVUC[0])

    def filtrar_bases_por_data(self) -> None:
        """Filtra bases por data"""

        for i in self.dicionarioDeBases:
            self.__dict__[i] = self.__dict__[i].loc[pd.to_datetime(self.__dict__[i]['Carimbo de data/hora'],yearfirst=False,dayfirst=True)\
             >= datetime.datetime(datetime.datetime.now().year,datetime.datetime.now().month,datetime.datetime.now().day)]

    def mapear_colunas_em_comum(self) -> list:
        """Mapea as colunas em comum das diferentes bases"""

        colunas = []

        for i in [self.baseInventario.columns,self.baseVUC.columns,self.baseDevolucao.columns,self.baseExpedicaoVUC.columns]:
            for j in i: colunas.append(j)

        colunas = list(set(colunas))
        
        return colunas

    def completar_colunas_em_falta(self) -> None:
        """Completa colunas de tabelas faltantes tomando como base as demais tabelas"""

        for i in self.mapear_colunas_em_comum():
            if i not in self.baseInventario.columns: self.baseInventario[i] = np.nan
            if i not in self.baseVUC.columns: self.baseVUC[i] = np.nan
            if i not in self.baseDevolucao.columns: self.baseDevolucao[i] = np.nan  
            if i not in self.baseExpedicaoVUC.columns: self.baseExpedicaoVUC[i] = np.nan

    def mostrar_head_das_tabelas(self) -> None:
        """Mostra head das tabelas"""

        print(self.baseConcatenada)

    def concatenar_tabelas(self) -> None:
        """Concatenar tabelas tratadas"""

        self.baseConcatenada = pd.concat([self.baseInventario,self.baseVUC,self.baseDevolucao,self.baseExpedicaoVUC])

    def preparar_tabela(self) -> object:

        """Prepara tabela para inserir dados na pasta do Excel"""

        self.mostrar_progresso(self.baixar_bases)
        self.mostrar_progresso(self.tratar_bases)
        self.mostrar_progresso(self.mapear_bases_a_manipular)
        self.mostrar_progresso(self.filtrar_bases_por_data)
        self.mostrar_progresso(self.completar_colunas_em_falta)
        self.mostrar_progresso(self.concatenar_tabelas)
        
        return self.baseConcatenada

    def mostrar_progresso(self, processo) -> None:
        """Mostra progresso do processo de preparo de tabelas"""
        
        processo()

        self.contagemProgresso += 1
        os.system('cls')
        barradeprogresso = '='*int(self.contagemProgresso*100/6)
        print(f'Excecutando função {self.contagemProgresso} de 6\n{self.contagemProgresso/6:.1%} |{barradeprogresso}|')

# comentário aleatório
