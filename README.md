# Assistente Inteligente de Seleção de Aeronaves Executivas

Sistema multiagente desenvolvido em Python para automatizar a análise de requisitos de clientes e auxiliar na seleção de aeronaves executivas, utilizando Inteligência Artificial e dados técnicos do *Business & Commercial Aviation (BCA)*.

O sistema integra quatro agentes especializados, coordenados por uma interface gráfica desenvolvida em PyQt6:

1. **Agente Vendedor:** interpreta documentos PDF/DOCX ou descrições textuais, identifica necessidades operacionais e gera um perfil estruturado do cliente em JSON.
2. **Agente Mercado:** extrai e organiza informações técnicas e comerciais das aeronaves presentes no catálogo BCA, utilizando processamento determinístico ou modelos de IA.
3. **Agente Engenheiro:** compara os requisitos da missão com as características das aeronaves, avaliando alcance, capacidade de passageiros, comprimento de pista e restrições operacionais.
4. **Agente Relatório:** consolida os resultados da análise técnica e gera automaticamente um relatório executivo em PDF, apresentando a recomendação preliminar e alternativas.

A **interface PyQt6** permite processar múltiplos clientes, executar agentes individualmente ou em sequência, acompanhar logs em tempo real e reutilizar resultados intermediários.

**Objetivo:** desenvolver uma solução de apoio à decisão que integre Inteligência Artificial, análise determinística e automação para tornar o processo de seleção de aeronaves executivas mais eficiente, padronizado e rastreável.
