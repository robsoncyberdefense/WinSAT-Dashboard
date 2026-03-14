# 📊 WinSAT Performance Dashboard

Transforme dados brutos de benchmark do Windows em relatórios profissionais de HTML automaticamente.

Uma ferramenta PowerShell desenvolvida para analistas de Cyber Security, SysAdmins e suporte de TI que precisam identificar gargalos de hardware (Disco, RAM, CPU) e provar se a lentidão do sistema é causada por limitações físicas ou configurações de software (como EDR/Antivírus).

## 🚀 Por que usar?

Muitas ferramentas de terceiros são pesadas, pagas ou exigem instalação. Este script utiliza o **WinSAT (Windows System Assessment Tool)**, uma ferramenta nativa e assinada pela Microsoft, para:

1.  Executar benchmarks oficiais de Disco, Memória, CPU e Gráficos.
2.  Extrair métricas críticas (Leitura Sequencial/Aleatória, Bandwidth, FPS do DWM).
3.  Gerar um **Dashboard HTML visual e interativo** pronto para apresentação a clientes ou gestão.
4.  Automatizar a defesa técnica: Prove com dados que o gargalo é o HD mecânico, não o antivírus.

## 🛠️ Requisitos

-   Windows 10 ou Windows 11.
-   PowerShell 5.1 ou superior (já vem instalado).
-   **Execução como Administrador** (Necessário para rodar o `winsat formal`).
-   Notebook conectado à tomada (O WinSAT bloqueia execução em bateria).
-   Ambiente Físico (O teste não roda em VMs ou sessões RDP remotas).

## ⚡ Como Usar

1.  Baixe este repositório clicando em **Code > Download ZIP** ou clone:
    ```bash
    git clone https://github.com/robsoncyberdefense/WinSAT-Dashboard.git
    ```

2.  Abra o **PowerShell como Administrador**.

3.  Navegue até a pasta onde baixou os arquivos:
    ```powershell
    cd caminho\para\WinSAT-Dashboard
    ```

4.  Execute o script diretamente da raiz:
    ```powershell
    .\GerarRelatorio_WinSAT.ps1
    ```
    *(Se der erro de execução, rode: `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` antes)*

5.  Aguarde ~2 minutos. O relatório será gerado em `C:\Relatorios\Relatorio_Completo_Performance.html` e aberto automaticamente no navegador.

## 📸 Exemplo de Saída

O script gera um relatório HTML semelhante a este:

![Exemplo do Dashboard](print_dashboard.png)

## 🔍 O que é analisado?

O script extrai dados dos arquivos XML gerados pelo WinSAT (`Formal.Assessment`, `Disk`, `Mem`, `DWM`):

| Componente | Métricas Extraídas | Importância para Segurança/TI |
| :--- | :--- | :--- |
| **Disco** | Leitura Sequencial, Aleatória, Modelo | Identifica se o I/O saturado causa travamento em varreduras de EDR. |
| **Memória** | Largura de Banda (Bandwidth) | Detecta se a RAM é insuficiente para multitarefa moderna. |
| **CPU** | Criptografia, Compressão | Avalia capacidade de processamento de heurísticas em tempo real. |
| **Gráficos** | FPS do DWM, Banda de Vídeo | Analisa a fluidez da interface (útil para troubleshooting de UI). |

## ⚠️ Limitações Conhecidas

-   **Virtualização:** O comando `winsat formal` não executa dentro de Máquinas Virtuais (VMware, VirtualBox, Hyper-V) ou ambientes Cloud PC (Windows 365), retornando Erro 13.
-   **Sessão Remota:** Não execute via RDP ou TeamViewer. O teste exige sessão local console (Erro 8/9).
-   **Energia:** Notebooks devem estar conectados à fonte de energia.

## 🤝 Contribuição

Sinta-se à vontade para abrir Issues ou Pull Requests se quiser adicionar novas métricas ou melhorar o layout do HTML.

## 📄 Licença

Este projeto está sob a licença MIT. Sinta-se livre para usar em seus ambientes corporativos.

---
**Autor:** Robson Nunes - Cyber Security
**LinkedIn:** https://www.linkedin.com/in/robsonsecurity
