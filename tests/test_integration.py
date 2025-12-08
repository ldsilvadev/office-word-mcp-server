"""
Testes de Integração para o MCP Server.
Testa o fluxo completo de criação e manipulação de documentos.

Execute com: pytest tests/test_integration.py -v
"""
import pytest
import os
import json
import asyncio
import tempfile
import shutil
from pathlib import Path


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


class TestIntegrationFlows:
    """Testes de integração para fluxos completos."""

    @pytest.mark.asyncio
    async def test_int001_create_document_flow(self, temp_dir):
        """INT-001: Fluxo completo de criação de documento."""
        from word_document_server.tools import document_tools, content_tools
        
        # 1. Criar documento
        doc_path = os.path.join(temp_dir, "integration_test.docx")
        result = await document_tools.create_document(doc_path, title="Teste Integração")
        assert "successfully" in result.lower()
        assert os.path.exists(doc_path)
        
        # 2. Adicionar conteúdo
        await content_tools.add_heading(doc_path, "Título Principal", 1)
        await content_tools.add_paragraph(doc_path, "Primeiro parágrafo do documento.")
        await content_tools.add_heading(doc_path, "Seção 2", 2)
        await content_tools.add_paragraph(doc_path, "Conteúdo da seção 2.")
        
        # 3. Verificar conteúdo
        text = await document_tools.get_document_text(doc_path)
        assert "Título Principal" in text
        assert "Primeiro parágrafo" in text
        assert "Seção 2" in text
        
        # 4. Obter estrutura
        outline = await document_tools.get_document_outline(doc_path)
        outline_data = json.loads(outline)
        assert isinstance(outline_data, (dict, list))

    @pytest.mark.asyncio
    async def test_int005_template_all_features(self, template_path, temp_dir):
        """INT-005: Template com todas as features."""
        if not os.path.exists(template_path):
            pytest.skip("Template não encontrado")
        
        from word_document_server.tools import document_tools
        
        output_path = os.path.join(temp_dir, "full_features.docx")
        
        # JSON completo com todas as features
        data = {
            "assunto": "Política Completa de Teste",
            "codigo": "POL-TEST-001",
            "departamento": "TI",
            "revisao": "01",
            "data_publicacao": "07/12/2025",
            "data_vigencia": "01/01/2026",
            "secao": [
                {
                    "titulo": "1. Objetivo",
                    "paragrafo": "Esta política estabelece as diretrizes para testes automatizados."
                },
                {
                    "titulo": "2. Escopo",
                    "paragrafo": "Aplica-se a todos os sistemas:\n- Sistema A\n- Sistema B\n- Sistema C"
                },
                {
                    "titulo": "3. Responsabilidades",
                    "paragrafo": "As responsabilidades são:\n1. Desenvolver testes\n2. Executar testes\n3. Reportar resultados",
                    "tabela_dinamica": [
                        {"cargo": "QA", "responsabilidade": "Criar testes", "prazo": "Semanal"},
                        {"cargo": "Dev", "responsabilidade": "Corrigir bugs", "prazo": "Diário"},
                        {"cargo": "PM", "responsabilidade": "Acompanhar", "prazo": "Mensal"}
                    ],
                    "paragrafo_pos_tabela": "Os prazos podem ser ajustados conforme necessidade do projeto."
                },
                {
                    "titulo": "4. Disposições Finais",
                    "paragrafo": "Esta política entra em vigor na data de sua publicação."
                }
            ]
        }
        
        result = await document_tools.fill_document_simple(
            template_path,
            output_path,
            json.dumps(data)
        )
        
        # Verificar que o documento foi criado
        assert os.path.exists(output_path)
        
        # Verificar conteúdo
        text = await document_tools.get_document_text(output_path)
        assert "Política Completa de Teste" in text or "POL-TEST-001" in text

    @pytest.mark.asyncio
    async def test_int006_template_empty_loop(self, template_path, temp_dir):
        """INT-006: Template com loop vazio."""
        if not os.path.exists(template_path):
            pytest.skip("Template não encontrado")
        
        from word_document_server.tools import document_tools
        
        output_path = os.path.join(temp_dir, "empty_loop.docx")
        
        data = {
            "assunto": "Documento Vazio",
            "codigo": "DOC-EMPTY",
            "secao": []  # Loop vazio
        }
        
        result = await document_tools.fill_document_simple(
            template_path,
            output_path,
            json.dumps(data)
        )
        
        # Documento deve ser criado mesmo com loop vazio
        assert os.path.exists(output_path)

    @pytest.mark.asyncio
    async def test_int007_template_empty_table(self, template_path, temp_dir):
        """INT-007: Template com tabela vazia."""
        if not os.path.exists(template_path):
            pytest.skip("Template não encontrado")
        
        from word_document_server.tools import document_tools
        
        output_path = os.path.join(temp_dir, "empty_table.docx")
        
        data = {
            "assunto": "Documento com Tabela Vazia",
            "secao": [
                {
                    "titulo": "Seção",
                    "paragrafo": "Texto",
                    "tabela_dinamica": []  # Tabela vazia
                }
            ]
        }
        
        result = await document_tools.fill_document_simple(
            template_path,
            output_path,
            json.dumps(data)
        )
        
        # Documento deve ser criado, placeholder removido
        assert os.path.exists(output_path)


class TestTableIntegration:
    """Testes de integração para tabelas."""

    @pytest.mark.asyncio
    async def test_table_full_workflow(self, temp_dir):
        """Fluxo completo de criação e formatação de tabela."""
        from word_document_server.tools import document_tools, content_tools, format_tools
        
        doc_path = os.path.join(temp_dir, "table_test.docx")
        
        # 1. Criar documento
        await document_tools.create_document(doc_path)
        
        # 2. Adicionar tabela
        data = [
            ["Nome", "Idade", "Cidade"],
            ["João", "30", "São Paulo"],
            ["Maria", "25", "Rio de Janeiro"],
            ["Pedro", "35", "Belo Horizonte"]
        ]
        await content_tools.add_table(doc_path, 4, 3, data)
        
        # 3. Formatar tabela
        await format_tools.highlight_table_header(doc_path, 0)
        await format_tools.apply_table_alternating_rows(doc_path, 0)
        
        # 4. Verificar
        assert os.path.exists(doc_path)
        text = await document_tools.get_document_text(doc_path)
        assert "João" in text
        assert "Maria" in text


class TestSearchReplaceIntegration:
    """Testes de integração para busca e substituição."""

    @pytest.mark.asyncio
    async def test_search_replace_workflow(self, temp_dir):
        """Fluxo de busca e substituição."""
        from word_document_server.tools import document_tools, content_tools
        
        doc_path = os.path.join(temp_dir, "search_test.docx")
        
        # 1. Criar documento com texto
        await document_tools.create_document(doc_path)
        await content_tools.add_paragraph(doc_path, "O valor antigo é 100.")
        await content_tools.add_paragraph(doc_path, "Outro valor antigo aqui.")
        
        # 2. Substituir
        result = await content_tools.search_and_replace(doc_path, "antigo", "novo")
        
        # 3. Verificar
        text = await document_tools.get_document_text(doc_path)
        assert "novo" in text
        assert "antigo" not in text


class TestEdgeCases:
    """Testes de casos extremos."""

    @pytest.mark.asyncio
    async def test_edg002_special_characters(self, temp_dir):
        """EDG-002: Caracteres especiais."""
        from word_document_server.tools import document_tools, content_tools
        
        doc_path = os.path.join(temp_dir, "special_chars.docx")
        
        await document_tools.create_document(doc_path)
        
        # Texto com caracteres especiais
        special_text = "Texto com émojis 🎉 e acentuação: ção, ñ, ü, ß, 中文"
        await content_tools.add_paragraph(doc_path, special_text)
        
        # Verificar
        text = await document_tools.get_document_text(doc_path)
        assert "ção" in text
        # Emojis podem ou não ser preservados dependendo da implementação

    @pytest.mark.asyncio
    async def test_edg005_section_without_paragraph(self, template_path, temp_dir):
        """EDG-005: Seção sem parágrafo."""
        if not os.path.exists(template_path):
            pytest.skip("Template não encontrado")
        
        from word_document_server.tools import document_tools
        
        output_path = os.path.join(temp_dir, "no_paragraph.docx")
        
        data = {
            "assunto": "Teste",
            "secao": [
                {"titulo": "Título Apenas", "paragrafo": ""}
            ]
        }
        
        result = await document_tools.fill_document_simple(
            template_path,
            output_path,
            json.dumps(data)
        )
        
        # Não deve dar erro
        assert os.path.exists(output_path)


class TestErrorHandling:
    """Testes de tratamento de erros."""

    @pytest.mark.asyncio
    async def test_err005_invalid_json(self, template_path, temp_dir):
        """ERR-005: JSON malformado."""
        if not os.path.exists(template_path):
            pytest.skip("Template não encontrado")
        
        from word_document_server.tools import document_tools
        
        output_path = os.path.join(temp_dir, "error_test.docx")
        
        result = await document_tools.fill_document_simple(
            template_path,
            output_path,
            "{ invalid json }"
        )
        
        assert "error" in result.lower()

    @pytest.mark.asyncio
    async def test_nonexistent_file(self):
        """Operação em arquivo inexistente."""
        from word_document_server.tools import content_tools
        
        result = await content_tools.add_paragraph(
            "/caminho/inexistente/arquivo.docx",
            "Texto"
        )
        
        assert "not exist" in result.lower() or "error" in result.lower()


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
