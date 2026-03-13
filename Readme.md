
# 📄 UC EXPORT 


Software desenvolvido para **Exportar informações de XML** de **Primários, Secundários e Cromatografia** e gerar automaticamente:

- 📑 **Preenchimento da planilha de Cálculo de Incerteza - Gás, HP Flare, Óleo**

---

## ⚙️ Funcionalidades Principais

- 📄 Leitura dos certificado em XML  
- 🔎 Extração de informações pertinentes para elaboração do Cálculo de Incerteza
  
- ⚠️ Verificação de divergências de:
  - CARREGAMENTO CORRETO DO XML
  - VERIFICAÇÃO DA INSERÇÃO NA ORDEM CORRETA DOS INSTRUMENTOS DIFERENCIAIS DE PRESSÃO (ALTA, MEDIA E BAIXA)  
- 🔄 ATUALIZAÇÃO AUTOMÁTICA DOS CAMPOS DA PLANILHA  
- ✍️ INSERÇÃO MANUAL DOS DADOS DE OPERAÇÃO CASO NECESSÁRIO  
- 📑 PREENCHIMENTO AUTOMÁTICO E SALVA NA PLANILHA.

---

## 🧰 Suporte a Diferentes Tipos de XML de Instrumentos

- PT / PIT  
- DPT  
- TT / TIT  
- TE
- PRIMÁRIOS
- PLACA DE ORIFÍCIO  

---

## 📁 Estrutura dos Arquivos Necessários

```text
/CI_Export_app
│── ci_export_app.exe               → Executável
│── interface.xlsx           → Interface do software
│── loader                   → Módulos de leitura do XML
├── logo/                    → Imagens da interface
├── services/                → Regras de validação + módulos de identificação planilha e execução de fluxo
├── utils/                   → funções úteis para o software
├── writers/                 → Módulo com funções que escrevem os dados na planilha
└── gui/                     → Interface gráfica (Tkinter)
```

▶️ Como usar

![Tela inicial do ACs Generator](logo\CI Export.png)

1-Abra o software (executável).
2-Clique em Selecionar CI (selecione o arquivo em excel).
3-O programa:

- Abre a planilha
- Pergunta sobre qual tipo de XML quer importar 
- Exibe divergências caso econtre

-Gera automaticamente:

    - Prenchimento da planilha de Cálculo de incerteza selecionada
    

4-Planilha é salva na mesma pasta em que é selecionada.

OBS: XML's compatíveis no padrão Petrobras V3.0, todos os xml's podem ser gerados através do AC's app