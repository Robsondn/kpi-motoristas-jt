# KPI Motoristas â€” J&T Express SP

Dashboard web de acompanhamento de KPIs operacionais dos motoristas da J&T Express Filial SP. Alimentado por robÃ´s de coleta automÃ¡tica de dados do sistema BRDrive, com atualizaÃ§Ã£o em tempo real via servidor local.

---

## Funcionalidades

- Painel de KPIs por motorista: volumes, produtividade e desempenho
- Coleta automÃ¡tica de dados via dois robÃ´s integrados ao BRDrive (sistema de rastreamento)
- Servidor local em Python com compartilhamento na rede interna (acesso multi-usuÃ¡rio)
- Dashboard HTML completo, sem dependÃªncia de frameworks externos
- Agendamento automÃ¡tico dos robÃ´s via PowerShell
- Log de execuÃ§Ã£o com histÃ³rico de atualizaÃ§Ãµes

## Tecnologias

| Componente | Tecnologia |
|------------|-----------|
| Dashboard | HTML5, CSS3, JavaScript vanilla |
| Servidor | Python (http.server) |
| Coleta de dados | Python + BRDrive API |
| AutomaÃ§Ã£o | PowerShell (rodar_robos.ps1) |
| Dados | Excel (.xlsx) como base intermediÃ¡ria |

## Estrutura

```
Controle KPIs Motoristas/
â”œâ”€â”€ Dashboard_KPI_JT.html     # Dashboard principal
â”œâ”€â”€ index.html                # PÃ¡gina de entrada
â”œâ”€â”€ servidor_dashboard.py     # Servidor HTTP local com auto-open
â”œâ”€â”€ Robo controle brdrive1.py # RobÃ´ 1: coleta dados do BRDrive
â”œâ”€â”€ Robo dashboard brdrive2.py# RobÃ´ 2: processa e atualiza dashboard
â”œâ”€â”€ rodar_robos.ps1           # Script PowerShell para executar robÃ´s
â”œâ”€â”€ Iniciar Servidor Dashboard.bat  # Atalho para iniciar servidor
â””â”€â”€ backup kips Motoristas.xlsx     # Base de dados histÃ³rica
```

## Como usar

1. Execute `Iniciar Servidor Dashboard.bat` â€” abre automaticamente no navegador
2. Para atualizar dados: execute `rodar_robos.ps1` no PowerShell
3. Dashboard fica acessÃ­vel na rede local pelo IP da mÃ¡quina

## Contexto

Ferramenta criada para dar visibilidade rÃ¡pida ao desempenho individual dos motoristas durante a operaÃ§Ã£o diÃ¡ria, integrando dados do sistema de rastreamento BRDrive sem necessidade de acesso manual ao sistema.

---

*Desenvolvido por Robson Noberto â€” Analista de Dados | J&T Express Filial SP*

