# Office-Word-MCP-Server

Servidor MCP (Model Context Protocol) para criar e manipular documentos Microsoft Word através de assistentes de IA.

## 🚀 Início Rápido

```bash
# Criar ambiente virtual
python -m venv venv

# Ativar ambiente virtual
source venv/Scripts/activate

# Instalar dependências
pip install -e .

# Executar servidor
python -m word_document_server.main
```

## 📋 Funcionalidades

### Gerenciamento de Documentos

- Criar, copiar e converter documentos Word
- Extrair texto e analisar estrutura
- Preencher templates com dados dinâmicos
- Converter para PDF

### Criação de Conteúdo

- Adicionar títulos, parágrafos e quebras de página
- Inserir tabelas e imagens
- Criar listas numeradas e com marcadores
- Adicionar notas de rodapé

### Formatação

- Formatar texto (negrito, itálico, cores, fontes)
- Estilizar tabelas (bordas, cores, mesclagem de células)
- Buscar e substituir texto
- Aplicar estilos personalizados

### Recursos Avançados

- Proteção com senha
- Extração de comentários
- Manipulação de células de tabela
- Alinhamento e espaçamento

## 💾 Instalação

### Requisitos

- Python 3.8 ou superior
- pip

### Instalação Básica

```bash
# Clonar repositório
git clone https://github.com/ldsilvadev/office-word-mcp-server.git
cd office-word-mcp-server

# Instalar dependências
pip install -r requirements.txt
```

## ⚙️ Configuração com IDEs que suportão MCP

Adicione ao arquivo JSON de configuração da IDE:

**Configuração:**

```json
{
  "mcpServers": {
    "word-document-server": {
      "command": "python",
      "args": ["/caminho/para/word_mcp_server.py"]
    }
  }
}
```

Reinicie a sua IDE após salvar.

## 💬 Exemplos de Uso

Após configurar, você pode pedir a sua IDE:

- "Crie um documento chamado 'relatorio.docx'"
- "Adicione um título e três parágrafos"
- "Insira uma tabela 4x4 com dados de vendas"
- "Formate a palavra 'importante' em negrito e vermelho"
- "Substitua 'termo antigo' por 'termo novo'"
- "Adicione uma lista numerada com três itens"
- "Extraia todos os comentários do documento"
- "Preencha o template 'modelo.docx' com dados JSON"
- "Converta o documento para PDF"

## 🔧 Principais Funções

### Documentos

- `create_document()` - Criar documento
- `convert_to_pdf()` - Converter para PDF
- `copy_document()` - Copiar documento

### Conteúdo

- `add_heading()` - Adicionar título
- `add_paragraph()` - Adicionar parágrafo
- `add_table()` - Adicionar tabela
- `add_picture()` - Adicionar imagem

### Formatação

- `format_text()` - Formatar texto
- `format_table()` - Formatar tabela
- `search_and_replace()` - Buscar e substituir

### Cabeçalhos e Rodapés

- `get_header_text()` - Ler texto do cabeçalho
- `set_header_text()` - Definir texto do cabeçalho
- `get_footer_text()` - Ler texto do rodapé
- `set_footer_text()` - Definir texto do rodapé

### Templates

- `fill_document_template()` - Preencher com Jinja2
- `fill_document_simple()` - Substituição simples

## 🔍 Solução de Problemas

### Problemas Comuns

**Permissões:** Verifique se o servidor tem permissão de leitura/escrita nos arquivos.

**Imagens:** Use caminhos absolutos e formatos compatíveis (JPEG, PNG).

**Tabelas:** Use cores hexadecimais sem '#' (ex: "FF0000" para vermelho).

### Debug

Ative logs detalhados:

```bash
# Windows
set MCP_DEBUG=1

# Linux/macOS
export MCP_DEBUG=1
```

## 📄 Licença

MIT License - veja o arquivo LICENSE para detalhes.

## 🙏 Créditos

- [Model Context Protocol](https://modelcontextprotocol.io/)
- [python-docx](https://python-docx.readthedocs.io/)
- [FastMCP](https://github.com/modelcontextprotocol/python-sdk)

---

**Nota:** Este servidor manipula arquivos no seu sistema. Sempre verifique as operações antes de confirmar.
