# Automação da Semana Universitária da UnB 📘

Projeto desenvolvido para automatizar a **organização e geração de eBook** da Semana Universitária da Universidade de Brasília (UnB).

## 🎯 Objetivo
Automatizar o processamento de planilhas de eventos e gerar automaticamente um eBook padronizado em PowerPoint (A4), reduzindo o tempo manual e garantindo uniformidade visual e textual.

## ⚙️ Estrutura do projeto
- **limpador.py** → Limpeza e padronização da planilha:
  - Remoção de duplicatas.
  - Correção de acentuação e formatação (*title case* e *sentence case*).
  - Normalização de textos com expressões regulares.
- **Código_SemUni.py** → Geração automática do eBook:
  - Criação de slides A4 com `python-pptx`.
  - Adaptação dinâmica das caixas de texto.
  - Aplicação da identidade visual da Semana Universitária.

## 🧠 Tecnologias utilizadas
- Python 3  
- pandas  
- python-pptx  
- re (expressões regulares)  
- unicodedata  

## 📘 Resultado
Geração de um **EBook completo e padronizado** com os projetos do evento, diretamente a partir da planilha da UnB.

## 📸 Exemplo de resultado
![Exemplo do eBook](<img width="218" height="320" alt="Captura de tela 2025-11-10 153522" src="https://github.com/user-attachments/assets/8fb98f73-9349-4cb3-b3f3-30991fd93204" />)


## 🧑‍💻 Autor
[Rodrigo Lobato] – [LinkedIn](https://www.linkedin.com/in/rblobato/)

