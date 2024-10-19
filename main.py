import json
import numpy as np
import pandas as pd
from math import ceil


# Exceção personalizada para lidar com erros na planilha
class PlanilhaException(Exception):
    pass


class PlanilhaPresencas:
    def __init__(self):
        self.qtd_eventos_extras = 0  # Inicializa a contagem de eventos extras

        # Inicializa a lista de petianos e conta a quantidade
        self.petianos = self.definirPetianos()  # Obtém os nomes dos petianos
        self.qtd_petianos = len(self.petianos)  # Conta quantos petianos existem
        self.qtd_faltas = [0] * self.qtd_petianos  # Inicializa contagem de faltas

        # Inicializa os eventos e suas informações
        self.lista_eventos = []  # Lista para armazenar os nomes dos eventos registrados
        self.num_eventos_especificos = []  # Lista para armazenar a contagem de eventos específicos
        self.nome_eventos, self.datas, self.extras = self.definirEventos()  # Obtém informações dos eventos
        self.qtd_eventos = len(self.nome_eventos)  # Conta quantos eventos existem
        self.qtd_eventos_petiano = [self.qtd_eventos] * self.qtd_petianos  # Inicializa contagem de eventos por petiano

        # Cria a matriz da planilha
        self.mat_planilha = self.definirMatrizPlanilha()  # Inicializa a matriz para registrar as presenças
        self.ajustarAtividades()  # Preenche a matriz com atividades padrão e extras

        # Executa o método principal para registrar presenças
        self.executar()

    def definirPetianos(self):
        # Lê o arquivo JSON e extrai os nomes dos petianos
        with open('petianos.json', 'r', encoding='utf-8') as arquivo:
            dados = json.load(arquivo)

        petianos = dados['nomes']  # Extrai os nomes dos petianos
        return petianos

    def registrarEvento(self, evento, extra):
        # Registra um evento e aumenta a contagem se já existir
        registrado = False

        for i in range(len(self.lista_eventos)):
            if self.lista_eventos[i] == evento:  # Verifica se o evento já está registrado
                indice = i
                registrado = True
                break

        if not registrado:  # Se não está registrado, adiciona
            self.lista_eventos.append(evento)
            self.num_eventos_especificos.append(1)  # Inicia a contagem desse evento

            if extra:  # Se é um evento extra, incrementa a contagem
                self.qtd_eventos_extras += 1
        else:
            self.num_eventos_especificos[indice] += 1  # Incrementa a contagem do evento já registrado

    def definirEventos(self):
        # Inicializa listas para armazenar informações dos eventos
        nome_eventos = []
        datas = []
        extras = []

        # Lê o arquivo JSON e extrai os dados dos eventos
        with open('eventos.json', 'r', encoding='utf-8') as file:
            eventos = json.load(file)

            for evento in eventos['eventos']:
                nome_eventos.append(evento['nome'])  # Adiciona o nome do evento
                datas.append(evento['data'])  # Adiciona a data do evento
                extras.append(evento['extra'])  # Adiciona se é evento extra
                self.registrarEvento(evento['nome'], evento['extra'])  # Registra o evento

        return nome_eventos, datas, extras

    def definirMatrizPlanilha(self):
        # Define o número de colunas adicionais com base nos eventos
        if (len(self.lista_eventos) - self.qtd_eventos_extras > self.qtd_eventos_extras):
            num_col_add = len(self.lista_eventos) - self.qtd_eventos_extras
        else:
            num_col_add = self.qtd_eventos_extras

        # Cria uma matriz com a quantidade de petianos e 7 colunas
        mat_planilha = np.full((self.qtd_petianos + num_col_add + 3, 7), ' ', dtype=object)
        colunas = ['Atividades Padrão', 'Atividades Extras', 'Número de Faltas', 'Faltas Justificadas',
                   'Limites de Faltas', 'N° Excedente']  # Nomes das colunas

        for i in range(self.qtd_petianos + 1):
            for j in range(7):
                # Preenche a primeira coluna com os nomes dos petianos
                if j == 0:
                    if i == 0:
                        mat_planilha[i][j] = 'Petianos'  # Cabeçalho da coluna de petianos
                    else:
                        mat_planilha[i][j] = self.petianos[i - 1]  # Nome do petiano

                # Preenche a primeira linha com os nomes das colunas
                elif i == 0:
                    mat_planilha[i][j] = colunas[j - 1]  # Nome da coluna

                else:
                    mat_planilha[i][j] = '0'  # Inicializa outros valores como '0'

        return mat_planilha

    def procurarIndice(self, evento):
        # Procura o índice do evento na lista de eventos
        for i in range(len(self.lista_eventos)):
            if evento == self.lista_eventos[i]:
                return i  # Retorna o índice do evento encontrado

        return 0  # Retorna 0 se não encontrar

    def ajustarAtividades(self):
        # Ajusta a matriz para incluir as atividades padrão e extras
        self.mat_planilha[self.qtd_petianos + 2][0] = 'Atividades Padrão'
        self.mat_planilha[self.qtd_petianos + 2][1] = 'Atividades Extras'
        eventos_registrados = np.zeros(len(self.lista_eventos), dtype=bool)  # Array para controlar eventos registrados

        i = self.qtd_petianos + 3  # Índice para preencher atividades

        for k in range(self.qtd_eventos):
            if not self.extras[k]:  # Verifica se não é evento extra
                l = self.procurarIndice(self.nome_eventos[k])  # Encontra o índice do evento

                if not eventos_registrados[l]:  # Se o evento ainda não foi registrado
                    self.mat_planilha[i][0] = '{} ({})'.format(self.nome_eventos[k], self.num_eventos_especificos[
                        l])  # Preenche a atividade padrão
                    eventos_registrados[l] = True  # Marca como registrado
                    i += 1  # Incrementa o índice

        i = self.qtd_petianos + 3  # Reinicia o índice

        for k in range(self.qtd_eventos):
            if self.extras[k]:  # Verifica se é evento extra
                l = self.procurarIndice(self.nome_eventos[k])  # Encontra o índice do evento

                if not eventos_registrados[l]:  # Se o evento ainda não foi registrado
                    self.mat_planilha[i][1] = '{} ({})'.format(self.nome_eventos[k], self.num_eventos_especificos[
                        l])  # Preenche a atividade extra
                    eventos_registrados[l] = True  # Marca como registrado
                    i += 1  # Incrementa o índice

    def ajustarPresencas(self, opcao, i, j):
        # Ajusta as contagens de presenças e faltas com base na opção selecionada
        if opcao == 3:  # Opção para evento não obrigatório
            self.qtd_eventos_petiano[i] -= 1
            return

        if self.extras[j - 1]:  # Se é uma atividade extra
            aux = int(self.mat_planilha[i + 1][1])  # Atualiza atividades extras
            aux += 1
            self.mat_planilha[i + 1][1] = str(aux)  # Converte para string e armazena
        else:
            aux = int(self.mat_planilha[i + 1][2])  # Atualiza faltas
            aux += 1
            self.mat_planilha[i + 1][2] = str(aux)

        if opcao == 0:  # Faltou
            self.qtd_faltas[i] += 1
            aux = int(self.mat_planilha[i + 1][3])  # Atualiza faltas justificadas
            aux += 1
            self.mat_planilha[i + 1][3] = str(aux)

        elif opcao == 2:  # Falta justificada
            aux = int(self.mat_planilha[i + 1][3])  # Atualiza faltas justificadas
            aux += 1
            self.mat_planilha[i + 1][3] = str(aux)
            aux = int(self.mat_planilha[i + 1][4])  # Atualiza limites de faltas
            aux += 1
            self.mat_planilha[i + 1][4] = str(aux)

    def atualizarLimiteFaltas(self, i):
        # Calcula e atualiza os limites de faltas para um petiano
        limite = self.qtd_eventos_petiano[i] * 0.25  # Limite de faltas
        excedente = ceil(self.qtd_faltas[i] - limite)  # Faltas além do limite

        if excedente < 0:
            excedente = 0  # Ajusta se excedente for negativo

        self.mat_planilha[i + 1][5] = str(limite)  # Atualiza limite na matriz
        self.mat_planilha[i + 1][6] = str(excedente)  # Atualiza excedente na matriz

    def executar(self):
        # Loop para processar cada evento e registrar as presenças
        for j in range(self.qtd_eventos):
            print("\n========== {} ({}) ==========".format(self.nome_eventos[j], self.datas[j]))
            print("Digite 0 para faltou, 1 para presente, 2 para falta justificada ou 3 para não obrigatória:\n")
            i = 0

            while i < self.qtd_petianos:
                try:
                    # Solicita a opção do usuário para cada petiano
                    opcao = int(input('{}: '.format(self.petianos[i])))

                    # Verifica se a opção é válida
                    if opcao < 0 or opcao > 3:
                        raise PlanilhaException()  # Lança uma exceção se a opção não for válida

                    # Chama o método para ajustar as presenças
                    self.ajustarPresencas(opcao, i, j)
                    i += 1  # Incrementa o índice
                except PlanilhaException as e:
                    print('\nINSIRA UM NÚMERO ENTRE 0 E 3!\n')  # Mensagem de erro
                except ValueError:
                    print("\nErro: Por favor, insira um número inteiro.\n")  # Mensagem de erro

        # Atualiza limites de faltas para cada petiano após o registro de presenças
        for i in range(self.qtd_petianos):
            self.atualizarLimiteFaltas(i)

        # Converte a matriz para um DataFrame e salva como arquivo Excel
        df = pd.DataFrame(self.mat_planilha.tolist())
        df.to_excel('Planilha_de_Presencas.xlsx', index=False, header=False)  # Salva a planilha em Excel


# Inicializa a classe para executar o programa
PlanilhaPresencas()
