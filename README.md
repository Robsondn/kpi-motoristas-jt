# KPI Motoristas — J&T Express SP

Dashboard web de acompanhamento de KPIs operacionais dos motoristas da J&T Express Filial SP. Alimentado por robôs de coleta automática de dados do sistema BRDrive, com atualização em tempo real via servidor local.

---

## Funcionalidades

- Painel de KPIs por motorista: volumes, produtividade e desempenho
- Coleta automática de dados via dois robôs integrados ao BRDrive (sistema de rastreamento)
- Servidor local em Python com compartilhamento na rede interna (acesso multi-usuário)
- Dashboard HTML completo, sem dependência de frameworks externos
- Agendamento automático dos robôs via PowerShell
- Log de execução com histórico de atualizações

## Tecnologias

| Componente | Tecnologia |
|------------|-----------|
| Dashboard | HTML5, CSS3, JavaScript vanilla |
| Servidor | Python (http.server) |
| Coleta de dados | Python + BRDrive API |
| Automação | PowerShell (rodar_robos.ps1) |
| Dados | Excel (.xlsx) como base intermediária |

## Estrutura

```
Controle KPIs Motoristas/
├── Dashboard_KPI_JT.html     # Dashboard principal
├── index.html                # Página de entrada
├── servidor_dashboard.py     # Servidor HTTP local com auto-open
├── Robo controle brdrive1.py # Robô 1: coleta dados do BRDrive
├── Robo dashboard brdrive2.py# Robô 2: processa e atualiza dashboard
├── rodar_robos.ps1           # Script PowerShell para executar robôs
├── Iniciar Servidor Dashboard.bat  # Atalho para iniciar servidor
└── backup kips Motoristas.xlsx     # Base de dados histórica
```

## Como usar

1. Execute `Iniciar Servidor Dashboard.bat` — abre automaticamente no navegador
2. Para atualizar dados: execute `rodar_robos.ps1` no PowerShell
3. Dashboard fica acessível na rede local pelo IP da máquina

## Contexto

Ferramenta criada para dar visibilidade rápida ao desempenho individual dos motoristas durante a operação diária, integrando dados do sistema de rastreamento BRDrive sem necessidade de acesso manual ao sistema.

---

*Desenvolvido por Robson Noberto — Analista de Processos | J&T Express Filial SP*
