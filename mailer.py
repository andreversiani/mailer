from tkinter import *
from tkinter.filedialog import askopenfilename
import re

import win32com.client as win32


class Equipamento:

    def __init__(self, tensao, type, qtd, descricao, et=None):
        self.et = et
        self.get_et()
        self.tensao = tensao
        self.type = type
        self.qtd = qtd
        self.descricao = descricao

    def get_et(self):
        if self.et != None:
            window = Tk()
            window.withdraw()
            filename = askopenfilename()
            self.et = filename


class Emailer:
    def __init__(self, receiver, subject, body, attachments=None):
        self.receiver = receiver
        self.subject = subject
        self.body = body
        self.attachments = attachments

    def prepare_emails(self):
        outlook = win32.Dispatch('outlook.application')
        email = outlook.CreateItem(0)
        email.Display()
        bodystart = re.search("<body.*?>", email.HTMLBody)
        email.HTMLBody = re.sub(bodystart.group(), bodystart.group() +
                                self.body, email.HTMLBody)
        email.To = self.receiver
        email.Subject = self.subject
        if self.attachments:
            email.Attachments.Add(self.attachments)
        email.CC = 'comercialenergia@visionsistemas.com.br'


class Fornecedor:
    def __init__(self, information):
        self.information = information


dict = {

    'WEG': {
        'AT': {
            'disj': True,
            'sec': False
        },
        'MT': {
            'disj': False,
            'sec': False
        },
        'contatos':
        [
            {
                'nome': 'Luiz Fernando',
                'email': 'andreversiani01@gmail.com',
                'tratamento': 'Prezado'
            },
            {
                'nome': 'Maria Fernanda',
                'email': 'lumini2@hotmail.com.br',
                'tratamento': 'Prezada'
            }
        ]
    },

    'GE': {
        'AT': {
            'disj': True,
            'sec': True
        },
        'MT': {
            'disj': True,
            'sec': True
        },
        'contatos':
        [
            {
                'nome': 'Celso',
                'email': 'andreversiani@visionsistemas.com.br',
                'tratamento': 'Prezado'
            },
        ]
    },
}


equipamentos = [Equipamento('AT', 'disj', 1, 'Disjuntor 138 kV, 2500A'), Equipamento(
    'MT', 'sec', 5, 'Seccionadora 13,8kV')]

resumo = {}
for fornecedor in dict:
    for equipamento in equipamentos:
        validator = dict[fornecedor][equipamento.tensao][equipamento.type]
        if validator:
            if fornecedor not in resumo:
                resumo.update(
                    {fornecedor: []})
            resumo[fornecedor].append(equipamento)


for fornecedor in resumo:
    contatos = dict[fornecedor]['contatos']
    if len(contatos) > 1:
        tratamento = 'Prezados'
        introducao = 'como estão?'
    if len(contatos) == 1:
        tratamento = contatos[0]['tratamento']
        introducao = contatos[0]['nome'] + ', '

    text = ''
    item = 1
    ets = []
    for eq in resumo[fornecedor]:
        if eq.et:
            ets.append(eq.et)

        text += f"""
        <tr>
            <td>{item}</td>
            <td>{eq.descricao}</td>
            <td>{eq.qtd}</td>
        </tr>
        """
        item += 1

        body = f"""
            <div>
                <p>{tratamento}, {introducao}</p>
                <p>Pedimos o orçamento dos seguintes items conforme a tebela:</p>
                <table>
                    <tr style="color:white; background-color:blue";>
                        <th">Item</th>
                        <th>Descrição</th>
                        <th>Qtd</th>
                    </tr>
                    {text}    
                </table>
            </div>
            """
    to = []
    for contato in contatos:
        to.append(contato['email'])
    to = ';'.join(to)
    ets = ';'.join(ets)

    emailer = Emailer(to, 'Cotação', body, ets)
    emailer.prepare_emails()
