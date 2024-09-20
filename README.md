
># SELECT LANGUAGE
<details>
<summary>ENGLISH VERSION</summary>

  - [English Version](#ENGLISH-VERSION)
    
  - [Description](#description)

  - [Additional Features](#additional-features)

  - [Project Structure](#project-structure)

  - [Depencencies](#dependencies)

  - [Configurations JSON](#configurations-json)

  - [How to Use](#how-to-use)

  - [Configuration](#configuration)

  - [Execution](#execution)

  - [Features](#features)

  - [Contribution](#contribution)

  - [License](#license)
</details>

<details>
<summary>PORTUGESE VERSION</summary>

  - [PORTUGUESE VERSION](#PORTUGUESE-VERSION)
    
  - [Descrição](#descrição)

  - [Funcionalidades Adicionais](#funcionalidades-adicionais)

  - [Estrutura do projeto](#estrutura-do-projeto)

  - [Dependências](#dependências)

  - [Configurações JSON](#configurações-json)

  - [Como Utilizar](#como-utilizar)

  - [Configurações](#configurações)

  - [Execução](#execução)

  - [Funcionalidades](#funcionalidades)

  - [Contribuição](#contribuição)

  - [Licença](#licença)
</details>

---

># ENGLISH VERSION

# OpenCV Report Automation System

This software is an enhanced version of the original repository **Automação_Relatorios_Por_Foto**. The key improvements implemented in this version include:

- **Image recognition library replacement**: The old `pyautogui` library ([pyautoguiDocumentation](https://pyautogui.readthedocs.io/en/latest/)) was replaced with `OpenCV` ([OpenCVDocumentation](https://docs.opencv.org/4.x/d6/d00/tutorial_py_root.html)), resulting in greater precision and faster image recognition.
  
- **Custom configuration fields**: The software now allows the customization of file names, folder paths, and other important configurations through JSON configuration files, making adjustments easier and avoiding **[Hard Coding](https://www.bloomtech.com/article/what-is-hard-coding)**.

- **UI/UX improvements**: The interface has been enhanced to provide a smoother and more intuitive user experience, utilizing customized components with `customtkinter` ([customtkinterDocumentation](https://customtkinter.tomschimansky.com/documentation/)).

## Description

This software was designed to automate repetitive and critical tasks related to the demand management system of Company X. The automation includes data collection and organization, report execution within the company's ERP system, and log sending. This improves operational efficiency by reducing the need for manual intervention in tasks like stock checking, purchase order tracking, and product movement control.

### Key Objectives:

- **Automated login and report generation**: The software automatically accesses the ERP via GO-Global, generates pre-configured reports, organizes the data, and sends the results to the appropriate location.

- **Execution schedule management**: Predefined schedules are used to execute tasks at specific times, as configured in the `agendamentos.json` file.

- **Log Sending**: The system sends logs documenting the automation process, allowing continuous monitoring of errors or failures.

This automation allows Company X's team to focus on more strategic tasks, improving decision-making with always-updated, easily accessible reports.

## Additional Features

- **Automatic log generation**: At the end of each execution, the software generates and sends detailed logs that help monitor the success of the execution or identify failures.

- **Customization through configurations**: Configuration files allow easy adaptation of the system for different scenarios and reports.

## Project Structure

The project contains the following main files:

- **main.py**: The main entry point of the program. Coordinates the execution of different system functionalities.
- **Enviar_Log.py**: Responsible for sending execution logs to an external system, facilitating process tracking.
- **RoboV4.py**: Contains the automation logic, including file manipulation and interaction with GO-Global software to access the ERP.
- **utils.py**: Contains auxiliary functions used by various other modules for common tasks like file manipulation and report generation.
- **widgetsCustom.py**: Defines custom graphical interface components used in the application.

## Dependencies

- Python 3.8 or higher
- Required modules:
  - `customtkinter`: Used to create the customized graphical interface.
  - `OpenCV`: For image manipulation and report processing.

## Configurations JSON

The software uses several configuration files to define important parameters:

- **configuracoes_login.json**: Contains login credentials to access the ERP via GO-Global.
- **configuracoes_software.json**: Defines the paths and arguments needed to access and run the remote ERP software.
- **configuracoes_nomes.json**: Maps file names and reports managed by the software.
- **configuracoes_telegram.json** Contains API credentiual to access Telegram API.
- **agendamentos.json**: Defines the automatic execution times for the reports.

## How to Use

1. **Install dependencies**: 
   Run the following command to install the necessary packages:
   ```bash
   pip install -r requirements.txt

## Configuration

2. Review the provided configuration files (configuracoes_login.json, configuracoes_software.json, configuracoes_nomes.json) and adjust them as needed.

## Execution
1. Run the main.py script to start the automation system:
   ```bash
   python main.py
   
## Features

- **Automated login and report generation:** Uses GO-Global to automatically connect to Company X's ERP and generate configured reports.
- **Log Sending:** The system sends execution logs for tracking.
- **Task Scheduling:** The software can automatically execute tasks at times configured in the agendamentos.json file.

## Contribution

If you would like to contribute to the project, follow these steps:

1. Fork the repository
2. Create a new branch: `git checkout -b feature/MyNewFeature.`
3. Make your changes and commit them: `git commit -m 'Adding new feature'.`
4. Push the branch: `git push origin feature/MyNewFeature.`
5. Submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

># PORTUGUESE VERSION

# Sistema de Automação de Relatórios OpenCV

Este software é uma versão aprimorada do repositório original **Automação_Relatorios_Por_Foto**. As principais melhorias implementadas nesta versão incluem: 

- **Troca de biblioteca de identificação de imagens**: A antiga biblioteca `pyautogui`([pyautoguiDocumentation](https://pyautogui.readthedocs.io/en/latest/)) foi substituída pela `OpenCV`([OpenCVDocumentation](https://docs.opencv.org/4.x/d6/d00/tutorial_py_root.html)), resultando em maior precisão e velocidade de reconhecimento de imagens.

- **Campos de configuração personalizados**: Agora, o software permite a customização de nomes de arquivos, caminhos de pastas e outras configurações importantes por meio de arquivos de configuração em JSON, facilitando ajustes, afim de evitar o **[Hard Coding](https://www.bloomtech.com/article/what-is-hard-coding)**.

- **Melhorias na UI/UX:** A interface foi aprimorada para proporcionar uma experiência de usuário mais fluida e intuitiva, utilizando componentes personalizados com `customtkinter`([customtkinterDocumentation](https://customtkinter.tomschimansky.com/documentation/)).

## Descrição

Este software foi projetado para automatizar tarefas repetitivas e críticas relacionadas ao sistema de gestão de demanda da empresa X. A automação inclui a coleta e organização de dados, execução de relatórios no ERP da empresa e envio de logs de execução. Isso melhora a eficiência operacional ao reduzir a necessidade de intervenção manual para tarefas como a verificação de estoque, o acompanhamento de pedidos de compra e o controle de movimentações de produtos.

### Objetivos do sistema:

- **Automação do login e geração de relatórios:** O software acessa automaticamente o ERP via GO-Global, gera relatórios pré-configurados, organiza os dados e envia os resultados para o local apropriado.

- **Gerenciamento de horários de execução:** Utiliza agendamentos predefinidos para executar tarefas em momentos específicos, conforme configurado no arquivo agendamentos.json.

- **Envio de Logs:** O sistema envia logs que documentam o processo de automação, permitindo um acompanhamento contínuo de erros ou falhas.

Essa automação permite à equipe da X concentrar-se em tarefas mais estratégicas, melhorando a tomada de decisões com relatórios sempre atualizados e facilmente acessíveis.

## Funcionalidades Adicionais

- **Geração automática de logs:** Ao final de cada execução, o software gera e envia logs detalhados que ajudam a monitorar o sucesso da execução ou a identificar falhas.

- **Personalização através de configurações:** Arquivos de configuração permitem a fácil adaptação do sistema para diferentes cenários e relatórios.


## Estrutura do Projeto

O projeto contém os seguintes arquivos principais:

- **main.py**: O ponto de entrada principal do programa. Coordena a execução das diferentes funcionalidades do sistema.
- **Enviar_Log.py**: Responsável por enviar logs de execução para um sistema externo, facilitando o acompanhamento do processo.
- **RoboV4.py**: Contém a lógica de automação, incluindo a manipulação de arquivos e interação com o software GO-Global para acessar o ERP.
- **utils.py**: Contém funções auxiliares utilizadas por vários outros módulos para tarefas comuns, como manipulação de arquivos e geração de relatórios.
- **widgetsCustom.py**: Define componentes personalizados da interface gráfica utilizados na aplicação.

## Dependências

- Python 3.8 ou superior
- Módulos necessários:
  - `customtkinter`: Utilizado para criar a interface gráfica customizada.
  - `OpenCV`: Para manipulação de imagens e relatórios.

## Configurações JSON

Existem vários arquivos de configuração que o software utiliza para definir parâmetros importantes:

- **configuracoes_login.json**: Contém as credenciais de login para acessar o ERP via GO-Global.
- **configuracoes_software.json**: Define os caminhos e argumentos necessários para acessar e executar o software remoto do ERP.
- **configuracoes_nomes.json**: Mapeia nomes de arquivos e relatórios que o software gerencia.
- **configuracoes_telegram.json** Contém as credencias de API para acessar a API do Telegram.
- **agendamentos.json**: Define os horários de execução automática para os relatórios.

## Como Utilizar

1. **Instale as dependências**: 
   Execute o seguinte comando para instalar os pacotes necessários:
   ```bash
   pip install -r requirements.txt
  
## Configurações

2. Verifique os arquivos de configuração fornecidos (configuracoes_login.json, configuracoes_software.json, configuracoes_nomes.json) e ajuste-os conforme necessário.

## Execução
1. Execute o script main.py para iniciar o sistema de automação:
   ```bash
   python main.py
   
## Funcionalidades

- **Automação de login e geração de relatórios:** Utiliza o GO-Global para se conectar automaticamente ao ERP da X e gerar os relatórios configurados.
- **Envio de Logs:** O sistema envia logs de execução para acompanhamento.
- **Agendamento de tarefas:** O software pode executar tarefas automaticamente em horários configurados no arquivo agendamentos.json.

## Contribuição

Se deseja contrinuir com o projeto, seiga os seguintes passos:

1. Faça um fork do repositorio
2. **Crie uma nova branch: git checkout -b feature/MinhaNovaFeature.
3. Faça suas alterações e commit: git commit -m 'Adicionando nova funcionalidade'.**
4. Faça push da branch: git push origin feature/MinhaNovaFeature.
5. Envie um Pull Request.

## Licença

**Este projeto está sob licença MIT - veja o arquivo LICENSE para mais detalhes.**
