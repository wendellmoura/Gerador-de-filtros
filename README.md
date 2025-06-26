# Excel/CSV para HTML Interativo

![Python Version](https://img.shields.io/badge/python-3.7%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)
![Status](https://img.shields.io/badge/status-stable-brightgreen)

**Converte arquivos Excel/CSV em páginas HTML interativas com recursos avançados de filtragem**

Uma ferramenta GUI que transforma seus dados tabulares em páginas web interativas com filtros avançados, tabelas dinâmicas e análise de dados integrada.

---

## ✨ Recursos Principais

- 🔍 **Filtros interativos** por coluna com busca
- 📊 **Tabelas dinâmicas** configuráveis
- 🔎 **Pesquisa global** com realce de resultados
- 📋 **Análise de duplicados** por coluna
- 📱 **Design responsivo** que se adapta a qualquer dispositivo
- ⚡ **Paginação automática** para grandes conjuntos de dados
- 🎨 **Personalização completa** da aparência

---

## 📥 Instalação

1. Clone o repositório:
```bash
git clone https://github.com/seu-usuario/excel-to-html-converter.git
cd excel-to-html-converter
```

2. Instale as dependências:
```bash
pip install -r requirements.txt
```

> **Dica:** O projeto requer Python 3.7 ou superior.

---

## 🚀 Como Usar

1. Execute a aplicação:
```bash
python main.py
```
Ou, se você salvou o código em outro arquivo:
```bash
python nome_do_arquivo.py
```

2. Na interface gráfica:
   - Clique em **Procurar** para selecionar um arquivo Excel (`.xlsx`, `.xls`) ou CSV.
   - Configure o nome do arquivo HTML de saída e o título da página.
   - Defina opções como linhas alternadas, rolagem horizontal e realce de pesquisa.
   - Visualize uma prévia dos dados carregados.
   - Exporte para HTML clicando em **Exportar HTML**.

3. Na aba **Tabela Dinâmica**:
   - Configure as colunas de linha, coluna (opcional), valor e a função de agregação (soma, média, contagem etc).
   - Clique em **Gerar Tabela Dinâmica** para visualizar o resultado com seus próprios filtros.

4. Abra o arquivo HTML gerado no seu navegador e explore:
   - Filtros avançados por coluna
   - Busca global com destaque
   - Tabela dinâmica, duplicados, e navegação responsiva

---

## 🖼️ Exemplo de Interface

![Interface do Exportador de Filtros](docs/interface_exemplo.png)

---

## 🛠️ Dependências

- Python >= 3.7
- pandas
- numpy
- tkinter

> Instale todas as dependências facilmente com:
> ```bash
> pip install pandas numpy
> ```

---

## 📄 Licença

Este projeto está licenciado sob a Licença MIT. Veja o arquivo [LICENSE](LICENSE) para mais detalhes.

---

## 👨‍💻 Autor

Desenvolvido por [Wendell Moura](https://github.com/wendellmoura)
