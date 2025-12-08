"""
Testes para as ferramentas de documento do MCP Server.
Execute com: pytest tests/test_document_tools.py -v
"""
import pytest
import os
import json
import asyncio
import tempfile
import shutil
from pathlib import Path

# Importar as ferramentas
from word_document_server.tools import document_tools, content_tools, format_tools


@pytest.fixture
def temp_dir():
    """Cria um diretório temporário para os testes."""
    temp = tempfile.mkdtemp()
    yield temp
    shutil.rmtree(temp, ignore_errors=True)


@pytest.fixture
def template_path():
    """Retorna o caminho do template de teste."""
    return os.path.join(os.path.dirname(__file__), "..", "templates", "template.docx")


@pytest.fixture
def sample_document(temp_dir):
    """Cria um documento de exemplo para testes."""
    doc_path = os.path.join(temp_dir, "test_doc.docx")
    asyncio.run(document_tools.create_document(doc_path, title="Test", author="Pytest"))
    return doc_path


class TestCreateDocument:
    """Testes para criação de documentos."""

    @pytest.mark.asyncio
    async def test_create_document_simple(self, temp_dir):
        """DOC-001: Criar documento vazio."""
        doc_path = os.path.join(temp_dir, "test.docx")
        result = await document_tools.create_document(doc_path)
        
        assert "successfully" in result.lower()
        assert os.path.exists(doc_path)

    @pytest.mark.asyncio
    async def test_create_document_with_metadata(self, temp_dir):
        """DOC-002: Criar documento com metadados."""
        doc_path = os.path.join(temp_dir, "test_meta.docx")
        result = await document_tools.create_document(
            doc_path, 
            title="Título Teste", 
            author="Autor Teste"
        )
        
        assert "successfully" in result.lower()
        assert os.path.exists(doc_path)

    @pytest.mark.asyncio
    async def test_create_document_auto_extension(self, temp_dir):
        """DOC-003: Criar documento sem extensão."""
        doc_path = os.path.join(temp_dir, "test_no_ext")
        result = await document_tools.create_document(doc_path)
        
        assert "successfully" in result.lower()
        assert os.path.exists(doc_path + ".docx")


class TestFillTemplate:
    """Testes para preenchimento de templates."""

    @pytest.mark.asyncio
    async def test_fill_template_simple(self, template_path, temp_dir):
        """TPL-001: Preencher template simples."""
        if not os.path.exists(template_path):
            pytest.skip("Template não encontrado")
        
        output_path = os.path.join(temp_dir, "output.docx")
        data = {"assunto": "Teste Simples", "codigo": "TST-001"}
        
        result = await document_tools.fill_document_simple(
            template_path, 
            output_path, 
            json.dumps(data)
        )
        
        assert os.path.exists(output_path)

    @pytest.mark.asyncio
    async def test_fill_template_with_loop(self, template_path, temp_dir):
        """TPL-002: Preencher template com loop."""
        if not os.path.exists(template_path):
            pytest.skip("Template não encontrado")
        
        output_path = os.path.join(temp_dir, "output_loop.docx")
        data = {
            "assunto": "Teste Loop",
            "codigo": "TST-002",
            "secao": [
                {"titulo": "Seção 1", "paragrafo": "Conteúdo 1"},
                {"titulo": "Seção 2", "paragrafo": "Conteúdo 2"}
            ]
        }
        
        result = await document_tools.fill_document_simple(
            template_path, 
            output_path, 
            json.dumps(data)
        )
        
        assert os.path.exists(output_path)

    @pytest.mark.asyncio
    async def test_fill_template_with_table(self, template_path, temp_dir):
        """TPL-003: Preencher template com tabela dinâmica."""
        if not os.path.exists(template_path):
            pytest.skip("Template não encontrado")
        
        output_path = os.path.join(temp_dir, "output_table.docx")
        data = {
            "assunto": "Teste Tabela",
            "codigo": "TST-003",
            "secao": [
                {
                    "titulo": "Seção com Tabela",
                    "paragrafo": "Texto antes da tabela",
                    "tabela_dinamica": [
                        {"coluna1": "valor1", "coluna2": "valor2"},
                        {"coluna1": "valor3", "coluna2": "valor4"}
                    ]
                }
            ]
        }
        
        result = await document_tools.fill_document_simple(
            template_path, 
            output_path, 
            json.dumps(data)
        )
        
        assert os.path.exists(output_path)

    @pytest.mark.asyncio
    async def test_fill_template_with_post_table_paragraph(self, template_path, temp_dir):
        """TPL-004: Preencher template com parágrafo pós-tabela."""
        if not os.path.exists(template_path):
            pytest.skip("Template não encontrado")
        
        output_path = os.path.join(temp_dir, "output_post_table.docx")
        data = {
            "assunto": "Teste Pós-Tabela",
            "codigo": "TST-004",
            "secao": [
                {
                    "titulo": "Seção Completa",
                    "paragrafo": "Texto antes",
                    "tabela_dinamica": [
                        {"item": "A", "valor": "1"}
                    ],
                    "paragrafo_pos_tabela": "Este texto aparece após a tabela."
                }
            ]
        }
        
        result = await document_tools.fill_document_simple(
            template_path, 
            output_path, 
            json.dumps(data)
        )
        
        assert os.path.exists(output_path)

    @pytest.mark.asyncio
    async def test_fill_template_nonexistent(self, temp_dir):
        """TPL-005: Template inexistente."""
        output_path = os.path.join(temp_dir, "output.docx")
        
        result = await document_tools.fill_document_simple(
            "inexistente.docx", 
            output_path, 
            "{}"
        )
        
        assert "not exist" in result.lower() or "error" in result.lower()

    @pytest.mark.asyncio
    async def test_fill_template_invalid_json(self, template_path, temp_dir):
        """TPL-006: JSON inválido."""
        if not os.path.exists(template_path):
            pytest.skip("Template não encontrado")
        
        output_path = os.path.join(temp_dir, "output.docx")
        
        result = await document_tools.fill_document_simple(
            template_path, 
            output_path, 
            "invalid json {"
        )
        
        assert "error" in result.lower()

    @pytest.mark.asyncio
    async def test_fill_template_with_bullet_list(self, template_path, temp_dir):
        """TPL-007: Preencher com lista bullet."""
        if not os.path.exists(template_path):
            pytest.skip("Template não encontrado")
        
        output_path = os.path.join(temp_dir, "output_bullet.docx")
        data = {
            "assunto": "Teste Lista",
            "secao": [
                {
                    "titulo": "Lista Bullet",
                    "paragrafo": "- Item 1\n- Item 2\n- Item 3"
                }
            ]
        }
        
        result = await document_tools.fill_document_simple(
            template_path, 
            output_path, 
            json.dumps(data)
        )
        
        assert os.path.exists(output_path)

    @pytest.mark.asyncio
    async def test_fill_template_with_numbered_list(self, template_path, temp_dir):
        """TPL-008: Preencher com lista numerada."""
        if not os.path.exists(template_path):
            pytest.skip("Template não encontrado")
        
        output_path = os.path.join(temp_dir, "output_numbered.docx")
        data = {
            "assunto": "Teste Lista Numerada",
            "secao": [
                {
                    "titulo": "Lista Numerada",
                    "paragrafo": "1. Primeiro item\n2. Segundo item\n3. Terceiro item"
                }
            ]
        }
        
        result = await document_tools.fill_document_simple(
            template_path, 
            output_path, 
            json.dumps(data)
        )
        
        assert os.path.exists(output_path)


class TestParagraphOperations:
    """Testes para operações com parágrafos."""

    @pytest.mark.asyncio
    async def test_add_paragraph_simple(self, sample_document):
        """PAR-001: Adicionar parágrafo simples."""
        result = await content_tools.add_paragraph(sample_document, "Texto de teste")
        assert "added" in result.lower()

    @pytest.mark.asyncio
    async def test_add_paragraph_with_style(self, sample_document):
        """PAR-002: Adicionar parágrafo com estilo."""
        result = await content_tools.add_paragraph(
            sample_document, 
            "Título", 
            style="Heading 1"
        )
        # Pode falhar se o estilo não existir, mas não deve dar erro
        assert "added" in result.lower() or "not found" in result.lower()

    @pytest.mark.asyncio
    async def test_add_paragraph_with_formatting(self, sample_document):
        """PAR-003: Adicionar parágrafo com formatação."""
        result = await content_tools.add_paragraph(
            sample_document, 
            "Texto formatado",
            bold=True,
            font_size=14,
            color="FF0000"
        )
        assert "added" in result.lower()

    @pytest.mark.asyncio
    async def test_edit_paragraph_text(self, sample_document):
        """PAR-004: Editar texto de parágrafo."""
        # Primeiro adiciona um parágrafo
        await content_tools.add_paragraph(sample_document, "Texto original")
        
        # Depois edita
        result = await content_tools.edit_paragraph_text(
            sample_document, 
            0, 
            "Texto modificado"
        )
        assert "updated" in result.lower() or "success" in result.lower()

    @pytest.mark.asyncio
    async def test_edit_paragraph_invalid_index(self, sample_document):
        """PAR-005: Editar parágrafo inexistente."""
        result = await content_tools.edit_paragraph_text(
            sample_document, 
            999, 
            "Texto"
        )
        assert "invalid" in result.lower()

    @pytest.mark.asyncio
    async def test_delete_paragraph(self, sample_document):
        """PAR-006: Deletar parágrafo."""
        # Adiciona parágrafos
        await content_tools.add_paragraph(sample_document, "Parágrafo 1")
        await content_tools.add_paragraph(sample_document, "Parágrafo 2")
        
        # Deleta o primeiro
        result = await content_tools.delete_paragraph(sample_document, 0)
        assert "deleted" in result.lower() or "success" in result.lower()


class TestTableOperations:
    """Testes para operações com tabelas."""

    @pytest.mark.asyncio
    async def test_add_table_empty(self, sample_document):
        """TBL-001: Adicionar tabela vazia."""
        result = await content_tools.add_table(sample_document, 3, 3)
        assert "added" in result.lower()

    @pytest.mark.asyncio
    async def test_add_table_with_data(self, sample_document):
        """TBL-002: Adicionar tabela com dados."""
        data = [["A", "B"], ["C", "D"]]
        result = await content_tools.add_table(sample_document, 2, 2, data)
        assert "added" in result.lower()


class TestSearchReplace:
    """Testes para busca e substituição."""

    @pytest.mark.asyncio
    async def test_search_and_replace(self, sample_document):
        """SRC-001: Buscar e substituir texto."""
        # Adiciona texto
        await content_tools.add_paragraph(sample_document, "Texto antigo aqui")
        
        # Substitui
        result = await content_tools.search_and_replace(
            sample_document, 
            "antigo", 
            "novo"
        )
        assert "replaced" in result.lower() or "occurrence" in result.lower()

    @pytest.mark.asyncio
    async def test_search_not_found(self, sample_document):
        """SRC-002: Buscar texto inexistente."""
        result = await content_tools.search_and_replace(
            sample_document, 
            "texto_inexistente_xyz", 
            "novo"
        )
        assert "not found" in result.lower() or "no occurrence" in result.lower()


class TestHeadings:
    """Testes para cabeçalhos."""

    @pytest.mark.asyncio
    async def test_add_heading_level1(self, sample_document):
        """HDR-001: Adicionar heading nível 1."""
        result = await content_tools.add_heading(sample_document, "Título Principal", 1)
        assert "added" in result.lower()

    @pytest.mark.asyncio
    async def test_add_heading_level2(self, sample_document):
        """HDR-002: Adicionar heading nível 2."""
        result = await content_tools.add_heading(sample_document, "Subtítulo", 2)
        assert "added" in result.lower()

    @pytest.mark.asyncio
    async def test_add_heading_with_formatting(self, sample_document):
        """HDR-003: Heading com formatação."""
        result = await content_tools.add_heading(
            sample_document, 
            "Título Formatado", 
            1,
            bold=True,
            font_size=18
        )
        assert "added" in result.lower()


class TestFileOperations:
    """Testes para operações de arquivo."""

    @pytest.mark.asyncio
    async def test_copy_document(self, sample_document, temp_dir):
        """FIL-001: Copiar documento."""
        dest_path = os.path.join(temp_dir, "copy.docx")
        result = await document_tools.copy_document(sample_document, dest_path)
        
        assert os.path.exists(dest_path)

    @pytest.mark.asyncio
    async def test_list_documents(self, temp_dir):
        """FIL-002: Listar documentos."""
        # Cria alguns documentos
        await document_tools.create_document(os.path.join(temp_dir, "doc1.docx"))
        await document_tools.create_document(os.path.join(temp_dir, "doc2.docx"))
        
        result = await document_tools.list_available_documents(temp_dir)
        
        assert "doc1.docx" in result
        assert "doc2.docx" in result

    @pytest.mark.asyncio
    async def test_get_document_info(self, sample_document):
        """FIL-003: Obter info do documento."""
        result = await document_tools.get_document_info(sample_document)
        
        # Deve retornar JSON válido
        info = json.loads(result)
        assert isinstance(info, dict)

    @pytest.mark.asyncio
    async def test_get_document_text(self, sample_document):
        """FIL-004: Extrair texto."""
        # Adiciona texto
        await content_tools.add_paragraph(sample_document, "Texto para extrair")
        
        result = await document_tools.get_document_text(sample_document)
        
        assert "Texto para extrair" in result

    @pytest.mark.asyncio
    async def test_get_document_outline(self, sample_document):
        """FIL-005: Obter estrutura."""
        # Adiciona headings
        await content_tools.add_heading(sample_document, "Seção 1", 1)
        await content_tools.add_heading(sample_document, "Seção 2", 1)
        
        result = await document_tools.get_document_outline(sample_document)
        
        # Deve retornar JSON válido
        outline = json.loads(result)
        assert isinstance(outline, (dict, list))


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
