from flask import Flask
from flask_sqlalchemy import SQLAlchemy
import random
import string
import win32com.client as win32

outlook = win32.Dispatch('outlook.application')

app = Flask(__name__)
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = True
app.config['SQLALCHEMY_DATABASE_URI'] = 'mysql://root:@localhost/clientes'

db = SQLAlchemy(app)

class Usuario(db.Model):
    __tablename__ = 'usuario'
    id = db.Column(db.Integer, primary_key= True)
    nome = db.Column(db.String(50))
    email = db.Column(db.String(100))
    mensagem = db.Column(db.String(50)) #colocar boolean pra True or False


db.create_all()

"""
Populate the database with default user
"""
u1 = Usuario(
            nome="João",
            email="projeto_ic1@yopmail.com",
            mensagem="não enviada"
    )
    
u2 = Usuario(
            nome="Amanda",
            email="projeto_ic2@yopmail.com",
            mensagem="não enviada"
    )

u3 = Usuario(
            nome="Isadora",
            email="projeto_ic3@yopmail.com",
            mensagem="não enviada" #true
    )

db.session.add(u1)
db.session.add(u2)
db.session.add(u3)
db.session.commit()


lista_teste = list(Usuario.query)
print(lista_teste)

for Usuario.query in lista_teste:

    if Usuario.query.mensagem == 'não enviada':
        
        def senha_aleatoria(length):
            # combinação de letras maiúsculas e minúsculas
            resultado_str = ''.join(random.choice(string.ascii_letters) for i in range(length))
            return resultado_str

        senha = senha_aleatoria(8)
        nome = Usuario.query.nome
        login = Usuario.query.email

        mail = outlook.CreateItem(0)
        mail.To = Usuario.query.email
        mail.Subject = 'Cadastro finalizado!'
        mail.Body = '''
        Prezado/a {}, 
    
        Boas vindas! Seu cadastro foi finalizado, segue abaixo seu login e sua senha temporária:
    
        Login: {}
        Senha: {}
    
        Qualquer dúvida estou à disposição.
    
        At.te,
        Equipe Suporte
        '''.format(nome, login, senha)

        mail.Send()

        update_msg = Usuario.query
        update_msg.mensagem = 'Enviada!'
        db.session.commit()

