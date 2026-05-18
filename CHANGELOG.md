# Changelog

Todas as mudanças relevantes deste projeto serão documentadas aqui.
Formato baseado em [Keep a Changelog](https://keepachangelog.com/pt-BR/1.0.0/).

---

## [1.0.0] - 2026-05-15

### Added
- Lançamento inicial do projeto
- Dashboard com 4 páginas de análise financeira pessoal (Visão Geral, Lançamentos, Detalhamento, Registros)
- Pipeline completo Excel → Power Query → Power BI com atualização via refresh
- Sistema de macros VBA com fila automatizada — roda ao abrir o arquivo ou via botão
- Projeção automática de lançamentos fixos por frequência (mensal, semanal, quinzenal, bimestral, trimestral, semestral, anual)
- Geração automática de parcelamentos com rastreamento por ID único
- Cálculo automático de status de fatura por cartão de crédito
- Visuais 100% customizados em SVG (cards KPI, gauges, donuts) via DAX — sem visuais pagos
- Modelo star schema com tabela de datas e 128 medidas DAX organizadas por pasta