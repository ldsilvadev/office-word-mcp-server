# Casos de Teste - Office Word MCP Server

## 1. Variáveis Simples

### TC-001: Substituição de variável simples
**Objetivo:** Verificar substituição de placeholders básicos

**Template:**
```
{{assunto}}
{{codigo}}
```

**JSON de entrada:**
```json
{
  "assunto": "Política de Home Office",
  "codigo": "FIERGS-001"
}
```

**Resultado esperado:** Texto substituído corretamente mantendo formatação original

---

### TC-002: Variável em cabeçalho/rodapé
**Objetivo:** Verificar substituição em headers e footers

**Template:** Header com `{{departamento}}` e footer com `{{revisao}}`

**JSON de entrada:**
```json
{
  "departamento": "Recursos Humanos",
  "revisao": "1.0"
}
```

**Resultado esperado:** Variáveis substituídas no cabeçalho e rodapé

---

### TC-003: Variável dentro de tabela do template
**Objetivo:** Verificar substituição dentro de células de tabela

**Template:** Tabela com células contendo `{{assunto}}` e `{{codigo}}`

**JSON de entrada:**
```json
{
  "assunto": "Teste",
  "codigo": "001"
}
```

**Resultado esperado:** Valores substituídos dentro das células da tabela

---

## 2. Loop de Seções

### TC-010: Loop básico com título e parágrafo
**Objetivo:** Verificar expansão de loop com múltiplas seções

**Template:**
```
{{LOOP:secao}}
1    {{titulo}}
        {{paragrafo}}
```

**JSON de entrada:**
```json
{
  "secao": [
    {"titulo": "1. Objetivo", "paragrafo": "Esta política define..."},
    {"titulo": "2. Escopo", "paragrafo": "Aplica-se a todos..."},
    {"titulo": "3. Responsabilidades", "paragrafo": "O gestor deve..."}
  ]
}
```

**Resultado esperado:** 3 seções criadas com títulos e parágrafos correspondentes

---

### TC-011: Loop apenas com título (seON de en

**JSioom array vaz cmportamento co VerificarObjetivo:**
**vaziaão seçLoop com ### TC-012: 
---

erro
os, sem idxpandtítulos eApenas ** erado:o espResultad```

**
"}
  ]
} 2"Item": itulo
    {"t"},"Item 1ulo": "tit
    { [":ao
  "sec
{json```:**
e entrada*JSON d``

*
`itulo}}{tsecao}}
{OOP:
```
{{L:**lateemp

**Trafoder de parágaceholem plunciona sr que loop frificaivo:** Ve)
**Objetrafoágm par