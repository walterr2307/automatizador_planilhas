import numpy as np
import pandas as pd


class PlanilhaIndividual:
    def __init__(self, petianos, eventos, lista_eventos, eventos_presentes, qtd_faltas, qtd_faltas_just,
                 qtd_eventos_petiano, eventos_extra_petiano, datas):
        self.petianos = petianos
        self.qtd_petianos = len(petianos)
        self.eventos = eventos
        self.lista_eventos = lista_eventos
        self.eventos_presentes = eventos_presentes
        self.qtd_faltas = qtd_faltas
        self.qtd_faltas_just = qtd_faltas_just
        self.qtd_eventos_petiano = qtd_eventos_petiano
        self.eventos_extra_petiano = eventos_extra_petiano
        self.planilha = self.definirPlanilha()
        self.datas = datas
        self.addFaltas()
        self.addDatas()

        df = pd.DataFrame(self.planilha.tolist())
        df.to_excel('Planilha_Individual.xlsx', index=False, header=False)

    def definirPlanilha(self):
        planilha = np.full((7 + len(self.lista_eventos), self.qtd_petianos * 5 - 1), fill_value=' ', dtype=object)
        lista_linha0 = ['ATIVIDADES', 'DATAS', 'PRESENTE']

        for k in range(self.qtd_petianos):
            for i in range(len(self.lista_eventos) + 1):
                for j in range(4):

                    if i == 0:
                        if j == 0:
                            planilha[i][j + k * 5] = self.petianos[k]
                        else:
                            planilha[i][j + k * 5] = lista_linha0[j - 1]

                    elif j == 1:
                        planilha[i][j + k * 5] = self.lista_eventos[i - 1]

        return planilha

    def addFaltas(self):
        for k in range(self.qtd_petianos):
            self.planilha[-4][1 + k * 5] = 'Total de Atividades:'
            self.planilha[-3][1 + k * 5] = 'Atividades Extras:'
            self.planilha[-2][1 + k * 5] = 'Atividades Presentes:'
            self.planilha[-1][1 + k * 5] = 'Faltas:'

            self.planilha[-4][2 + k * 5] = str(self.qtd_eventos_petiano[k])
            self.planilha[-3][2 + k * 5] = str(self.eventos_extra_petiano[k])
            self.planilha[-2][2 + k * 5] = str(self.qtd_eventos_petiano[k] - self.qtd_faltas[k])
            self.planilha[-1][2 + k * 5] = str(self.qtd_faltas[k])

    def addDatas(self):
        for k in range(self.qtd_petianos):
            for i in range(len(self.lista_eventos)):
                self.addDatasAtvds(k, i, self.lista_eventos[i])

    def addDatasAtvds(self, k, i, evento):
        datas = self.ajustarStrData(evento, k)
        self.planilha[i + 1][2 + k * 5] = datas
        datas = self.ajustarDatasPresentes(evento, k)
        self.planilha[i + 1][3 + k * 5] = datas

    def ajustarStrData(self, evento, k):
        string = ''

        for j in range(len(self.eventos)):
            if evento == self.eventos[j] and self.eventos_presentes[k][j] < 3:
                string += self.datas[j] + '\n'

        return string

    def ajustarDatasPresentes(self, evento, k):
        string = ''

        for j in range(len(self.eventos)):
            if evento == self.eventos[j] and self.eventos_presentes[k][j] == 1:
                string += self.datas[j] + '\n'

        return string
