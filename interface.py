import os
import customtkinter as ctk
from tkinter import filedialog, messagebox
from gerador import gerar_documento

PASTA_ATUAL = os.path.dirname(os.path.abspath(__file__))
CAMINHO_TEMPLATE = os.path.join(PASTA_ATUAL, 'modelo_template.docx')

ctk.set_appearance_mode ('light')
ctk.set_default_color_theme ('blue')

class App (ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title ('Gerador de Atestado Psicológico')
        self.geometry ('650x750')

        self.frame_scroll = ctk.CTkScrollableFrame(self,label_text='Dados do candidato')
        self.frame_scroll.pack(fill = 'both', expand= True, padx= 15, pady = 15)

        self.campos = {}

        self._montar_secao_dados_pessoais()
        self._montar_secao_avaliacao()
        self._montar_secao_aspectos_gerais()
        self._montar_secao_resultado()

        self.botao_gerar = ctk.CTkButton(
            self,
            text = "GERAR DOC", 
            height=40,
            font= ctk.CTkFont(size= 15, weight='bold'),
            command=self.ao_clicar_gerar, 
        )
        self.botao_gerar.pack(pady=(0,15))

    def _titulo_secao(self, texto):
        label = ctk.CTkLabel(
            self.frame_scroll,
            text=texto,
            font = ctk.CTkFont(size=16, weight= 'bold')
        )
        label.pack(anchor= 'w', pady = (15,5))

    def _campo_entry (self, chave, rotulo, placeholder= ''):
        ctk.CTkLabel(self.frame_scroll, text = rotulo).pack (anchor= 'w')
        entry = ctk.CTkEntry(self.frame_scroll, placeholder_text=placeholder)
        entry.pack(fill='x', pady= (0,8))
        self.campos[chave] = entry 

    def _campo_opcoes (self,chave,rotulo,opcoes):
        ctk.CTkLabel(self.frame_scroll, text = rotulo).pack(anchor='w')
        variavel= ctk.StringVar(value=opcoes[0])
        menu = ctk.CTkOptionMenu(self.frame_scroll, values= opcoes, variable = variavel)
        menu.pack(fill = 'x', pady = (0,8))

        self.campos[chave] = variavel

    def _montar_secao_dados_pessoais (self):
        self._titulo_secao('Dados Pessoais')
        self._campo_entry('nome_completo', 'Nome completo')
        self._campo_entry('idade', 'Idade')
        self._campo_opcoes('sexo', 'Sexo', ['FEMININO', 'MASCULINO'])
        self._campo_entry ('estado_civil', 'Estado civil')

    def _montar_secao_avaliacao(self):
        self._titulo_secao('Dados da avaliação')        
        self._campo_entry('instituicao', 'Instituição atendida')
        self._campo_entry ('finalidade', 'Finalidade')
        self._campo_entry('cargo', 'Cargo')
        self._campo_entry ('funcao', 'Função avaliada', 'Ex: Estágio em Administração')
        self._campo_entry('escolaridade', 'Escolaridade')
        self._campo_entry ('curso', 'Curso')
        self._campo_entry ('data_avaliacao', 'Data da avaliação', 'dd/mm/aaaa')

    def _montar_secao_aspectos_gerais (self):
        self._titulo_secao('Aspectos gerais')

        self.var_tem_filhos = ctk.BooleanVar(value=False)
        checkbox = ctk.CTkCheckBox(
            self.frame_scroll,
            text = 'Candidato(a) tem filhos? ', 
            variable=self.var_tem_filhos,
        )
        checkbox.pack(anchor='w', pady = (0,8))
        self.campos ['tem_filhos'] = self.var_tem_filhos

        self._campo_entry ('qtd_filhos', 'Quantos filhos? (deixe em branco caso não tenha)')

        ctk.CTkLabel(self.frame_scroll, text= "Experiência profissional (texto livre)").pack(anchor='w')

        self.textbox_experiencia = ctk.CTkTextbox(self.frame_scroll, height=80)
        self.textbox_experiencia.pack (fill = 'x', pady= (0,8))

        self._campo_entry ('percentual_portugues', '% aproveitamento em portugês')
        self._campo_entry ('percentual_matematica', '% aproveitamento em Matemática')
        self._campo_entry ('percentual_raciocinio', '% aproveitamento em raciocínio Lógico')

    #Resultado Final

    def _montar_secao_resultado (self):
        self._titulo_secao('Resultado')
        self._campo_opcoes ('resultado', 'Resultado da avaliação', ['APTA', 'INAPTA', 'APTO', 'INAPTO'])
        self._campo_entry('data_assinatura', 'Data assinatura (por extenso)', "Ex: 17 de agosto de 2026")


    def coletar_dados(self) -> dict:
        dados = {}

        for chave, widget in self.campos.items ():
            dados[chave] = widget.get()

        nome_completo = dados.get ('nome_completo',' ').strip()
        dados['nome_primeiro'] = nome_completo.split(' ') [0] if nome_completo else ''

        dados ['experiencia'] = self.textbox_experiencia.get("1.0", 'end').strip()
        return dados


    def validar_dados (self, dados:dict) ->list:
        obrigatorios = [
            "nome_completo", "idade", "estado_civil", "instituicao",
            "finalidade", "cargo", "funcao", "escolaridade", "curso",
            "data_avaliacao", "percentual_portugues", "percentual_matematica",
            "percentual_raciocinio", "data_assinatura",
        ]
        faltando = [campo for campo in obrigatorios if not dados.get (campo,'').strip()]
        return faltando

    def ao_clicar_gerar (self):
        dados = self.coletar_dados()

        faltando = self.validar_dados(dados)
        if faltando:
            messagebox.showerror(
                'Campos obrigatórios não preenchidos', 
                'Preencha os seguintes campos antes de gerar o documento: \n\n' + '\n'.join(f'- {campo}'for campo in faltando),
            )    
            return

        nome_sugerido = f'Atestado_{dados['nome_primeiro']}.docx'
        caminho_saida = filedialog.asksaveasfilename(
            defaultextension='.docx',
            filetypes=[('Documento Word', '*.docx')],
            initialfile= nome_sugerido, 
            title= 'Salvar atestado como...', 
    )    

        if not caminho_saida:
            return

        try:
            gerar_documento(dados, CAMINHO_TEMPLATE, caminho_saida)
        except Exception as erro:
            messagebox.showerror ('Erro ao gerar documento', str (erro))

            return
        messagebox.showinfo(f'Documento gerado com sucesso em: \n {caminho_saida}')

if __name__== "__main__":
    App = App()
    App.mainloop()    