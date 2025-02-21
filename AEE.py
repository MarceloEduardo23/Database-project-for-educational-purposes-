from sqlalchemy import create_engine, Column, String, Integer, ForeignKey
from sqlalchemy.orm import sessionmaker, declarative_base
import os,time
import pandas as pd


db = create_engine("sqlite:///Lista_de_alunos_AEE.db")
Session = sessionmaker(bind=db)
session = Session()

Base = declarative_base()

class Aluno(Base):
    __tablename__ = "Alunos"
    num = Column("Nº", Integer, primary_key=True, autoincrement=True)
    nome = Column("Nome", String)
    idade = Column("Idade", Integer)
    deficiencia = Column("Deficiência", String)
    cpf = Column("CPF", String)
    numero = Column("Número", String)
    
    def __init__(self, nome, idade, deficiencia, cpf, numero):
        self.nome = nome
        self.idade = idade
        self.deficiencia = deficiencia
        self.cpf = cpf
        self.numero = numero
        
Base.metadata.create_all(bind=db)

while True:
    print("\033[33mO que deseja fazer? \033[0m")
    loop = int(input("\033[31m1\033[0m \033[33mAdicionar aluno\033[33m\n\033[31m2\033[0m \033[33mModificar tabela\033[33m\n\033[31m0\033[0m \033[33mSair do programa\033[33m\n"))
    if loop == 1:
        nome = input("\033[33mNome do aluno(a):\033[0m ")
        idade = int(input("\033[33mIdade do aluno(a):\033[0m "))
        deficiencia = input("\033[33mDeficiência do aluno(a):\033[0m ")
        cpf = input("\033[33mDigite o CPF do aluno(a) \033[31m(ex: 111.222.333-44)\033[0m\033[33m:\033[0m ")
        numero = input("\033[33mDigite o telefone para contato \033[31m(ex: (11) 9 1111-8888)\033[0m\033[33m:\033[0m ")
        
        usuario = Aluno(nome = nome, idade = idade, deficiencia = deficiencia, cpf = cpf, numero = numero)
        session.add(usuario)
        session.commit()
        print("\033[32mAluno(a) Cadastrado com Sucesso!\033[0m")
        time.sleep(2)
        
    elif loop == 2:
        pesquisa = input("Qual o nome do Aluno(a) que deseja modificar? ")
        alterar = session.query(Aluno).filter_by(nome = pesquisa).first()
        
        if alterar == None:
            print("\033[31mAluno(a) não encontrado!\33[0m]")
            print("dica: digite o nome exatamente como a planilha!")
            time.sleep(2)
        else:        
            opcao = int(input("Qual informação deseja alterar:\n1 - CPF\n2 - Telefone\n3 - Deficiência\n"))
            if opcao == 1:
                cpf = input("\033[33mDigite o CPF do aluno(a) \033[31m(ex: 111.222.333-44)\033[0m\033[33m:\033[0m ")
                alterar.cpf = cpf
                session.commit()
            if opcao == 2:
                telefone = input("\033[33mDigite o telefone do aluno(a) \033[31m(ex: (11) 9 1111-8888)\033[0m\033[33m:\033[0m ")
                alterar.numero = telefone
                session.commit()
            if opcao == 3:
                deficiencia = input("\033[33mDigite a deficiência do aluno(a):\033[0m ")
                alterar.deficiencia = deficiencia
                session.commit()
            
            print("\033[32mInformação alterada com Sucesso!\033[0m")
            time.sleep(2)
    else:
        print("\033[32mAlterações concluídas\033[0m")
        break
    
    os.system('cls' if os.name == 'nt' else 'clear')

resposta = int(input("\033[33mDigite \033[31m1\033[0m \033[33mpara salvar planilha ou \033[0m\033[31m0\033[0m \033[33mpara sair: \033[0m"))

if resposta == 1:
    # Conectar ao banco de dados SQLite
    conexao = db.connect()

    # Ler os dados da tabela "Alunos" para um DataFrame pandas
    df = pd.read_sql("SELECT * FROM Alunos", conexao)

    # Salvar os dados no arquivo Excel
    df.to_excel("Lista_de_alunos_AEE.xlsx", index=False, engine="openpyxl")

    # Fechar a conexão
    conexao.close()

    print("\033[32mExportação concluída! O arquivo 'Lista_de_alunos_AEE.xlsx' foi criado.\033[0m")
    