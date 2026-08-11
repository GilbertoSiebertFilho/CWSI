#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Conteudo de referencia usado para montar a planilha CWSI + Guia de Monitores.

Separado do construtor para que voce possa editar/ampliar o conteudo sem mexer
na logica de montagem do arquivo .xlsx.
"""

# =========================================================================== #
#  1. BASELINES NAO-ESTRESSADAS (Idso, 1982)                                  #
#     Tc - Ta = a + b * VPD      (VPD em kPa, dT em °C, b negativo)           #
#     ATENCAO: valores de literatura, obtidos em Phoenix/AZ (clima arido).    #
#     Servem como PONTO DE PARTIDA. Calibracao local e' obrigatoria.          #
# =========================================================================== #
BASELINES = [
    # (cultura, a, b, observacao)
    ("Alfafa (alfalfa)",              0.51, -1.92, "Idso (1982)"),
    ("Algodao",                       1.49, -2.09, "Idso (1982) - baseline classica"),
    ("Abobora (squash)",              3.85, -2.75, "Idso (1982)"),
    ("Batata",                        1.17, -1.83, "Idso (1982)"),
    ("Beterraba",                     5.16, -2.30, "Idso (1982)"),
    ("Cevada (antes do espigamento)", 2.01, -2.25, "Idso (1982)"),
    ("Feijao",                        2.91, -2.35, "Idso (1982)"),
    ("Feijao-caupi",                  1.32, -1.84, "Idso (1982)"),
    ("Girassol",                      2.79, -2.31, "Idso (1982)"),
    ("Milho",                         3.11, -1.97, "Idso (1982)"),
    ("Pepino",                        4.88, -2.52, "Idso (1982)"),
    ("Alface",                        4.18, -2.96, "Idso (1982)"),
    ("Soja",                          1.44, -1.34, "Idso (1982)"),
    ("Tomate",                        2.86, -1.96, "Idso (1982)"),
    ("Trigo (antes do espigamento)",  2.88, -2.11, "Idso (1982)"),
    ("Trigo (apos o espigamento)",    1.72, -1.23, "Idso (1982)"),
    ("PERSONALIZADO (calibracao local)", None, None,
     "Use os campos manuais em 01_Config. RECOMENDADO."),
]

# =========================================================================== #
#  2. CATALOGO DE MONITORES                                                   #
# =========================================================================== #
# (arquivo_icone, marca, monitor, formato_fisico, tamanho, versoes,
#  plataforma_web, telemetria, midia, formatos_aceitos, observacoes)
MONITORES = [
    ("jd_gs3_2630.png", "John Deere", "GreenStar 3 2630",
     "10.4\" 4:3, tela sensivel ao toque + teclas laterais", "10.4 pol",
     "GS3 3.x (ex.: 3.34.1349) — atualizacao pelo StellarSupport",
     "Operations Center (antigo MyJohnDeere) / Apex (descontinuado)",
     "JDLink (opcional, via MTG)", "Pen drive USB (FAT/FAT32)",
     "Shapefile (.shp/.shx/.dbf/.prj); arquivos de Setup do Apex/Ops Center",
     "Geracao anterior ao Gen 4. Estrutura de pastas propria: GS3_2630\\<Perfil>\\RCD."),

    ("jd_gen4.png", "John Deere", "Gen 4 (4240 / 4600 CommandCenter / 4640 Universal)",
     "10.1\" 16:9, tela sensivel ao toque", "10.1 pol",
     "Gen 4 OS 10.x e 11.x (ex.: 10.16.1400, 11.x)",
     "Operations Center", "JDLink / MTG — permite envio sem fio",
     "Pen drive USB (FAT32) e sem fio", "Shapefile em pasta Rx; perfis .zip; ISOXML (conforme versao)",
     "Prescricoes devem estar na pasta 'Rx' na RAIZ do pen drive. Exportacao: Setup -> JD4600, Trabalho -> JD-Data."),

    ("jd_g5.png", "John Deere", "G5 (G5e / G5 / G5Plus)",
     "12.8\" 16:9, tela sensivel ao toque", "12.8 pol",
     "G5 OS (linha sucessora do Gen 4 OS)",
     "Operations Center", "JDLink integrado (conforme configuracao)",
     "Pen drive USB (FAT32) e sem fio", "Shapefile em pasta Rx; ISOXML; transferencia sem fio",
     "Fluxo de arquivos semelhante ao Gen 4. Prioriza transferencia sem fio pelo Operations Center."),

    ("cih_afs_pro700.png", "Case IH", "AFS Pro 700",
     "10.4\" 4:3, encoder rotativo + teclas", "10.4 pol",
     "AFS Pro 700 software 28.x / 30.x / 32.x (varia por maquina)",
     "AFS Connect", "AFS Connect (conforme maquina)",
     "Pen drive USB — CNH recomenda o PN 84398840",
     "Shapefile (prescricao); linhas de guia Multiswath; importa linhas de displays concorrentes",
     "Equivalente ao New Holland IntelliView IV. Confirme o caminho de menu na versao instalada."),

    ("cih_afs_pro1200.png", "Case IH", "AFS Pro 1200",
     "12.1\" 16:10, tela sensivel ao toque", "12.1 pol",
     "AFS Pro 1200 software (linha nova, atualizacao pela concessionaria/AFS Connect)",
     "AFS Connect", "AFS Connect — envio sem fio de prescricoes",
     "Pen drive USB e sem fio", "Shapefile (pasta selecionada na importacao); ISOXML",
     "Equivalente ao New Holland IntelliView 12. Na importacao voce seleciona a PASTA que contem o shapefile."),

    ("nh_intelliview4.png", "New Holland", "IntelliView IV",
     "10.4\" 4:3, encoder rotativo + teclas", "10.4 pol",
     "IntelliView IV software (mesma base do AFS Pro 700)",
     "PLM Connect / MyPLM Connect", "PLM Connect (opcional)",
     "Pen drive USB", "Shapefile; linhas de guia; ISOXML (conforme versao)",
     "Irmao do Case IH AFS Pro 700 (plataforma CNH). Procedimentos praticamente identicos."),

    ("nh_intelliview12.png", "New Holland", "IntelliView 12",
     "12.1\" 16:10, tela sensivel ao toque", "12.1 pol",
     "IntelliView 12 software (mesma base do AFS Pro 1200)",
     "PLM Connect", "PLM Connect — envio sem fio",
     "Pen drive USB e sem fio", "Shapefile; ISOXML",
     "Irmao do Case IH AFS Pro 1200 (plataforma CNH)."),

    ("trimble_gfx750.png", "Trimble", "GFX-750",
     "10.1\" 16:9, teclas fisicas inferiores", "10.1 pol",
     "Precision-IQ (varias versoes) — atualizacao por pen drive",
     "Trimble Ag Software / PTx Trimble", "Modem opcional",
     "Pen drive USB (FAT32)", "Shapefile em AgData\\Prescriptions; pacotes .zip AgData; ISOXML",
     "Precision-IQ usa a pasta 'AgData'. Prescricoes shapefile em AgData\\Prescriptions."),

    ("trimble_gfx1060.png", "Trimble", "GFX-1060 / GFX-1260",
     "10\" e 12\" 16:9, tela sensivel ao toque", "10 e 12 pol",
     "Precision-IQ (versoes recentes)",
     "Trimble Ag Software / PTx Trimble", "Modem integrado (conforme configuracao)",
     "Pen drive USB e sem fio", "Shapefile em AgData\\Prescriptions; .zip AgData; ISOXML",
     "Mesma logica de pastas do GFX-750."),

    ("trimble_tmx2050.png", "Trimble", "TMX-2050",
     "12.1\" 16:10, tela sensivel ao toque", "12.1 pol",
     "Precision-IQ / FmX Plus (conforme configuracao)",
     "Trimble Ag Software", "Modem opcional",
     "Pen drive USB", "ISOXML em pasta TASKDATA; AgData; AgGPS (FmX/FmX+)",
     "Aceita TASKDATA (ISOXML). Pacotes antigos usam a pasta AgGPS."),

    ("agleader_incommand.png", "Ag Leader", "InCommand 800 / InCommand 1200",
     "8\" e 12.1\", tela sensivel ao toque", "8 e 12.1 pol",
     "Firmware InCommand 1.x a 6.x (ex.: 3.5, 6.x)",
     "AgFiniti", "AgFiniti (sem fio, conforme licenca)",
     "Pen drive USB e AgFiniti", ".agsetup (configuracao: talhoes, limites, linhas de guia, prescricoes); .agdata (dados registrados); .pat (padroes entre geracoes)",
     ".agsetup so' troca entre grupos de display compativeis (Integra/Versa entre si; InCommand entre si). Para mover padroes entre geracoes, use .pat."),

    ("raven_viper4.png", "Raven", "Viper 4 / Viper 4+",
     "12.1\", tela sensivel ao toque", "12.1 pol",
     "ROS (Raven Operating Software) 3.x / 4.x",
     "Slingshot", "Slingshot Field Hub",
     "Pen drive USB e Slingshot", "Shapefile; pacotes de trabalho .zip; ISOXML (conforme versao)",
     "Estrutura de pastas gerada pelo proprio Viper ao exportar. Confirme na versao do ROS."),

    ("topcon_x35.png", "Topcon", "X35 / XD+ (Horizon)",
     "12.1\", tela sensivel ao toque, ISOBUS UT/TC-BAS/TC-GEO/TC-SC", "12.1 pol",
     "Horizon 5.x / 6.x",
     "Topcon Agriculture Platform (TAP)", "TAP / modem opcional",
     "Pen drive USB (FAT32)", "ISOXML (TASKDATA.XML + .BIN) na pasta TASKDATA; shapefile (conforme versao)",
     "Base dos terminais Valtra/Massey Ferguson em varias configuracoes. Padrao ISOBUS."),

    ("fendt_varioterminal.png", "AGCO (Fendt / Valtra / Massey Ferguson)",
     "Varioterminal / FendtONE / terminal ISOBUS AGCO",
     "10.4\" e 12\", ISOBUS", "10.4 e 12 pol",
     "Varia por linha (Varioterminal 10.4, FendtONE)",
     "AGCO Cloud / Fendt Task Doc / agrirouter", "TaskDoc Pro / VarioDoc Pro",
     "Pen drive USB (FAT32) e nuvem", "ISOXML (TASKDATA.XML + .BIN) na pasta TASKDATA",
     "100% ISOBUS (TC-BAS e TC-GEO). O arquivo .zip do escritorio contem TASKDATA.XML e arquivos .BIN."),
]

# =========================================================================== #
#  3. LISTAS DE SELECAO                                                       #
# =========================================================================== #
OBJETIVOS = [
    "Importar prescricao (taxa variavel)",
    "Importar linhas de guia (AB / curvas)",
    "Importar limites e cadastro (cliente/fazenda/talhao)",
    "Exportar dados de trabalho (as-applied / colheita)",
    "Exportar linhas de guia (AB)",
    "Atualizar software do monitor",
]

ORIGENS = [
    "Pen drive USB",
    "Nuvem / plataforma do fabricante",
]

TIPOS_EQUIPAMENTO = [
    "Trator",
    "Colheitadeira",
    "Pulverizador autopropelido",
    "Plantadeira / Semeadora",
    "Universal / portatil (retrofit)",
]

# Versoes por monitor. A primeira entrada e' sempre "Todas as versoes".
VERSOES = {
    "GreenStar 3 2630": ["Todas as versoes", "GS3 3.x"],
    "Gen 4 (4240 / 4600 CommandCenter / 4640 Universal)":
        ["Todas as versoes", "Gen 4 OS 10.x", "Gen 4 OS 11.x"],
    "G5 (G5e / G5 / G5Plus)": ["Todas as versoes", "G5 OS"],
    "AFS Pro 700": ["Todas as versoes", "Software 28.x", "Software 30.x ou superior"],
    "AFS Pro 1200": ["Todas as versoes"],
    "IntelliView IV": ["Todas as versoes"],
    "IntelliView 12": ["Todas as versoes"],
    "GFX-750": ["Todas as versoes", "Precision-IQ"],
    "GFX-1060 / GFX-1260": ["Todas as versoes", "Precision-IQ"],
    "TMX-2050": ["Todas as versoes", "Precision-IQ", "FmX Plus (AgGPS)"],
    "InCommand 800 / InCommand 1200":
        ["Todas as versoes", "Firmware 1.x a 3.x", "Firmware 4.x ou superior"],
    "Viper 4 / Viper 4+": ["Todas as versoes", "ROS 3.x", "ROS 4.x"],
    "X35 / XD+ (Horizon)": ["Todas as versoes", "Horizon 5.x", "Horizon 6.x"],
    "Varioterminal / FendtONE / terminal ISOBUS AGCO":
        ["Todas as versoes", "Varioterminal", "FendtONE"],
}

# =========================================================================== #
#  4. BASE DE PROCEDIMENTOS                                                   #
# =========================================================================== #
# Cada item: dicionario com marca, monitor, versao, objetivo, origem,
# formato, pasta, fs (sistema de arquivos), passos (lista), cuidados,
# erros, confianca, fonte.
#
# CONFIANCA:
#   "Verificado"  -> caminho de pasta / formato confirmado em fonte citada.
#   "Confirmar"   -> estrutura correta, mas o nome exato do menu muda entre
#                    versoes; confirme na maquina antes de publicar.
# =========================================================================== #

P = []


def add(**kw):
    P.append(kw)


# --------------------------------------------------------- JOHN DEERE ------
add(
    marca="John Deere", monitor="Gen 4 (4240 / 4600 CommandCenter / 4640 Universal)",
    versao="Todas as versoes",
    objetivo="Importar prescricao (taxa variavel)", origem="Pen drive USB",
    formato="Shapefile completo: .shp + .shx + .dbf + .prj (mesmo nome-base)",
    pasta="Rx\\  (na RAIZ do pen drive — nao pode estar dentro de outra pasta)",
    fs="FAT32 (particao MBR)",
    passos=[
        "Formate o pen drive em FAT32 com particao MBR.",
        "Na RAIZ do pen drive crie uma pasta chamada exatamente 'Rx'.",
        "Copie para dentro de 'Rx' os 4 arquivos do shapefile (.shp, .shx, .dbf, .prj) com o mesmo nome-base.",
        "Confira antes de ir a campo: geometria de POLIGONO, coordenadas WGS84 e pelo menos uma coluna NUMERICA com a taxa.",
        "Com a maquina ligada, insira o pen drive na porta USB do monitor.",
        "Toque no botao de Menu (canto inferior esquerdo da tela).",
        "Abra o 'Gerenciador de Arquivos' (File Manager).",
        "Selecione 'Importar Dados', escolha o pen drive como origem e marque a prescricao.",
        "Confirme a importacao e aguarde a barra de progresso terminar.",
        "Em Configuracao do Trabalho > Talhao > Prescricao, selecione o arquivo, a COLUNA DE TAXA e a UNIDADE; confira o mapa antes de iniciar.",
    ],
    cuidados="Nomes de arquivo sem acentos, sem espacos e curtos. Nao remova o pen drive durante a transferencia.",
    erros="Faltar o .prj (o monitor perde a projecao); shapefile em UTM em vez de WGS84; coluna de taxa gravada como TEXTO; pasta 'Rx' dentro de outra pasta; geometria de linha/ponto em vez de poligono.",
    confianca="Verificado",
    fonte="Deere StellarSupport (Gen 4 File Manager) + guia de formatacao de prescricao VRAFY",
)

add(
    marca="John Deere", monitor="Gen 4 (4240 / 4600 CommandCenter / 4640 Universal)",
    versao="Todas as versoes",
    objetivo="Exportar dados de trabalho (as-applied / colheita)", origem="Pen drive USB",
    formato="Pacote de dados do monitor (pasta gerada automaticamente)",
    pasta="Dados de trabalho -> JD-Data\\   |   Dados de configuracao (Setup) -> JD4600\\",
    fs="FAT32 (particao MBR)",
    passos=[
        "Encerre ou pause o trabalho para garantir que tudo foi gravado.",
        "Use UM pen drive por monitor (nao misture dados de maquinas diferentes no mesmo pen drive).",
        "Insira o pen drive na porta USB do monitor.",
        "Menu > Gerenciador de Arquivos (File Manager).",
        "Selecione 'Exportar Dados'.",
        "Escolha 'Dados de Trabalho' (Work Data) — o monitor grava na pasta 'JD-Data'.",
        "Se precisar tambem do cadastro, escolha 'Dados de Configuracao' (Setup) — grava na pasta 'JD4600'.",
        "Aguarde a barra de progresso concluir antes de remover o pen drive.",
        "No escritorio, importe a pasta no Operations Center ou no software de gestao.",
    ],
    cuidados="Exportar com o trabalho aberto pode gerar arquivo incompleto. Sempre aguarde a confirmacao na tela.",
    erros="Remover o pen drive cedo demais; usar o mesmo pen drive em varios monitores e sobrescrever pastas; formatar em exFAT/NTFS.",
    confianca="Verificado",
    fonte="Deere StellarSupport — Generation 4 Displays (release notes / File Manager)",
)

add(
    marca="John Deere", monitor="Gen 4 (4240 / 4600 CommandCenter / 4640 Universal)",
    versao="Todas as versoes",
    objetivo="Importar linhas de guia (AB / curvas)", origem="Pen drive USB",
    formato="Arquivo de Setup exportado do Operations Center (pacote de configuracao)",
    pasta="JD4600\\  (pasta de Setup gerada pelo Operations Center / exportacao do monitor)",
    fs="FAT32 (particao MBR)",
    passos=[
        "No Operations Center, selecione cliente/fazenda/talhao e as linhas de guia desejadas.",
        "Gere o arquivo de configuracao (Setup) para display Gen 4 e baixe para o PC.",
        "Copie a pasta gerada para a RAIZ do pen drive, mantendo a estrutura original.",
        "Insira o pen drive no monitor.",
        "Menu > Gerenciador de Arquivos > Importar Dados.",
        "Selecione o pacote de Setup e confirme.",
        "Escolha se deseja SUBSTITUIR ou MESCLAR com os dados ja existentes no monitor.",
        "Abra Configuracao do Trabalho > Guia (Guidance) e selecione a linha AB importada.",
        "Confirme o nome da linha e o talhao antes de iniciar o trabalho.",
    ],
    cuidados="A opcao 'Substituir' apaga cadastros do monitor. Em duvida, escolha 'Mesclar'.",
    erros="Renomear ou reorganizar as pastas do pacote (o monitor deixa de reconhecer); importar Setup de outra geracao de display.",
    confianca="Confirmar",
    fonte="Deere Operations Center / StellarSupport — confirmar nome do menu na versao instalada",
)

add(
    marca="John Deere", monitor="Gen 4 (4240 / 4600 CommandCenter / 4640 Universal)",
    versao="Todas as versoes",
    objetivo="Importar prescricao (taxa variavel)", origem="Nuvem / plataforma do fabricante",
    formato="Prescricao publicada no Operations Center (envio sem fio via JDLink)",
    pasta="Nao se aplica — transferencia sem fio",
    fs="Nao se aplica",
    passos=[
        "Confirme que a maquina tem MTG/JDLink ativo e aparece como conectada no Operations Center.",
        "No Operations Center, abra 'Analise/Arquivos' e carregue (upload) a prescricao.",
        "Associe a prescricao ao cliente, fazenda e talhao corretos.",
        "Use 'Enviar para a maquina' (Send to Machine) e selecione o equipamento de destino.",
        "No monitor, aceite a notificacao de transferencia recebida.",
        "Menu > Gerenciador de Arquivos para confirmar que o arquivo chegou.",
        "Em Configuracao do Trabalho > Prescricao, selecione o arquivo, a coluna de taxa e a unidade.",
        "Confira o mapa e as taxas na tela antes de iniciar.",
    ],
    cuidados="Sem sinal de celular no talhao a transferencia nao completa. Leve sempre uma copia em pen drive como plano B.",
    erros="Maquina nao pareada a organizacao correta; usuario sem permissao de envio; prescricao associada ao talhao errado.",
    confianca="Confirmar",
    fonte="John Deere Operations Center — confirmar fluxo na versao atual da plataforma",
)

add(
    marca="John Deere", monitor="GreenStar 3 2630", versao="Todas as versoes",
    objetivo="Importar prescricao (taxa variavel)", origem="Pen drive USB",
    formato="Shapefile completo: .shp + .shx + .dbf + .prj",
    pasta="Rx\\  (na RAIZ do pen drive)",
    fs="FAT ou FAT32 (particao MBR)",
    passos=[
        "Formate o pen drive em FAT ou FAT32 (MBR). O 2630 e' exigente com o formato.",
        "Crie a pasta 'Rx' na RAIZ do pen drive e copie os 4 arquivos do shapefile.",
        "Insira o pen drive no monitor. Se nao reconhecer, teste a outra porta USB.",
        "Acesse o menu do GreenStar 3 e va em Configuracao do Trabalho.",
        "Selecione Cliente / Fazenda / Talhao antes de carregar a prescricao.",
        "Abra a aba de Taxa Variavel / Prescricao e escolha 'Carregar prescricao'.",
        "Selecione o arquivo, depois a coluna de taxa e a unidade de medida.",
        "Confira o mapa renderizado na tela e as taxas minima e maxima antes de iniciar.",
    ],
    cuidados="O 2630 e' sensivel a pen drives grandes ou em exFAT. Prefira pen drives de 4 a 16 GB em FAT32.",
    erros="Pen drive em exFAT/NTFS; falta do .prj; nome de arquivo longo ou com acento.",
    confianca="Verificado",
    fonte="John Deere GS3 2630 User Guide + material de campo (PremierCrop: Loading Prescriptions GS2/GS3)",
)

add(
    marca="John Deere", monitor="GreenStar 3 2630", versao="Todas as versoes",
    objetivo="Exportar dados de trabalho (as-applied / colheita)", origem="Pen drive USB",
    formato="Dados gravados (RCD) no formato do GreenStar 3",
    pasta="GS3_2630\\<NomeDoPerfil>\\RCD\\",
    fs="FAT ou FAT32 (particao MBR)",
    passos=[
        "Encerre o trabalho no monitor.",
        "Insira o pen drive na porta USB.",
        "Acesse Menu > Gerenciador de Dados (Data Manager).",
        "Selecione 'Copiar dados para pen drive' / exportacao de dados gravados.",
        "Escolha o perfil desejado e confirme.",
        "O monitor grava na estrutura GS3_2630\\<Perfil>\\RCD.",
        "Aguarde a conclusao e so' entao remova o pen drive.",
        "No PC, importe a pasta GS3_2630 completa no software de escritorio.",
    ],
    cuidados="Copie a pasta GS3_2630 INTEIRA para o PC. Copiar so' o conteudo do RCD quebra a leitura.",
    erros="Renomear a pasta do perfil; copiar arquivos soltos sem a estrutura de pastas.",
    confianca="Verificado",
    fonte="John Deere GS3 2630 User Guide (estrutura GS3_2630 > Perfil > RCD)",
)

add(
    marca="John Deere", monitor="G5 (G5e / G5 / G5Plus)", versao="Todas as versoes",
    objetivo="Importar prescricao (taxa variavel)", origem="Pen drive USB",
    formato="Shapefile completo: .shp + .shx + .dbf + .prj",
    pasta="Rx\\  (na RAIZ do pen drive)",
    fs="FAT32 (particao MBR)",
    passos=[
        "Formate o pen drive em FAT32 (MBR).",
        "Crie a pasta 'Rx' na RAIZ e copie os 4 arquivos do shapefile.",
        "Insira o pen drive na porta USB do monitor.",
        "Menu > Gerenciador de Arquivos > Importar Dados.",
        "Selecione o pen drive e a prescricao desejada; confirme.",
        "Em Configuracao do Trabalho > Talhao > Prescricao, selecione o arquivo importado.",
        "Defina a coluna de taxa e a unidade.",
        "Confira o mapa e as taxas antes de iniciar o trabalho.",
    ],
    cuidados="O G5 prioriza o envio sem fio pelo Operations Center; o pen drive continua valendo como alternativa.",
    erros="Mesmos erros do Gen 4: falta de .prj, projecao errada, coluna de taxa como texto.",
    confianca="Confirmar",
    fonte="Fluxo herdado do Gen 4 — confirmar nomes de menu na versao do G5 OS instalada",
)

# ------------------------------------------------------------ CASE IH ------
add(
    marca="Case IH", monitor="AFS Pro 700", versao="Todas as versoes",
    objetivo="Importar prescricao (taxa variavel)", origem="Pen drive USB",
    formato="Shapefile (.shp + .shx + .dbf + .prj)",
    pasta="Pasta a sua escolha no pen drive — voce a seleciona na tela de importacao",
    fs="FAT32",
    passos=[
        "Use um pen drive compativel — a CNH recomenda o PN 84398840.",
        "Formate em FAT32 e copie o shapefile completo (.shp, .shx, .dbf, .prj) para uma pasta de nome simples.",
        "Com a maquina ligada, insira o pen drive na porta USB do monitor.",
        "Acesse a area de gerenciamento de dados / importacao do AFS Pro 700.",
        "Escolha importar Prescricao (Shapefile).",
        "Navegue ate a pasta do pen drive e selecione o arquivo .shp.",
        "Selecione a coluna de taxa e a unidade de medida.",
        "Vincule a prescricao ao Talhao correto.",
        "Confira o mapa e as taxas na tela antes de iniciar a operacao.",
    ],
    cuidados="O caminho exato do menu muda entre versoes de software. Confirme na maquina e registre o caminho real na base.",
    erros="Pen drive nao reconhecido (formato errado); shapefile incompleto; coluna de taxa como texto.",
    confianca="Confirmar",
    fonte="Case IH — Shapefile (.shp) Import for the AFS Pro 700 (QRC oficial); confirmar caminho de menu na versao instalada",
)

add(
    marca="Case IH", monitor="AFS Pro 700", versao="Todas as versoes",
    objetivo="Importar linhas de guia (AB / curvas)", origem="Pen drive USB",
    formato="Linhas de guia Multiswath; o AFS Pro 700 tambem importa linhas geradas em monitores concorrentes",
    pasta="Pasta selecionada na tela de importacao",
    fs="FAT32",
    passos=[
        "Gere/exporte as linhas de guia no software de escritorio ou no monitor de origem.",
        "Copie os arquivos para uma pasta de nome simples no pen drive (FAT32).",
        "Insira o pen drive na porta USB do AFS Pro 700.",
        "Acesse a area de gerenciamento de dados / importacao.",
        "Selecione a importacao de linhas de guia (Multiswath / linhas de concorrentes).",
        "Navegue ate a pasta e selecione o arquivo.",
        "Vincule as linhas ao Talhao correto.",
        "Abra a tela de Guia (Guidance) e selecione a linha importada.",
        "Confirme a largura de trabalho e o alinhamento antes de iniciar.",
    ],
    cuidados="Ao importar linhas de outra marca, confira o sistema de coordenadas e o deslocamento (offset) em campo.",
    erros="Largura de trabalho diferente da usada na criacao da linha; linha vinculada ao talhao errado.",
    confianca="Confirmar",
    fonte="Case IH — AFS Pro 700 Display: Competitive Guidance Line Importing; confirmar na versao instalada",
)

add(
    marca="Case IH", monitor="AFS Pro 1200", versao="Todas as versoes",
    objetivo="Importar prescricao (taxa variavel)", origem="Pen drive USB",
    formato="Shapefile (.shp + .shx + .dbf + .prj)",
    pasta="Voce seleciona a PASTA que contem o shapefile na janela 'Select Import Source'",
    fs="FAT32",
    passos=[
        "Formate o pen drive em FAT32.",
        "Crie uma pasta de nome simples e copie para dentro dela os 4 arquivos do shapefile.",
        "Insira o pen drive na porta USB do AFS Pro 1200.",
        "Na tela inicial, abra a area de importacao de dados.",
        "Na janela 'Select Import Source' (Selecionar origem da importacao), navegue e selecione a PASTA que contem o shapefile.",
        "Pressione 'Select' (Selecionar) para continuar.",
        "Escolha a coluna de taxa e a unidade de medida.",
        "Vincule a prescricao ao Cliente / Fazenda / Talhao.",
        "Confira o mapa e as taxas antes de iniciar o plantio/aplicacao.",
    ],
    cuidados="Na tela de importacao voce seleciona a PASTA, e nao o arquivo .shp diretamente.",
    erros="Selecionar o arquivo em vez da pasta; shapefile incompleto; projecao diferente de WGS84.",
    confianca="Verificado",
    fonte="Case IH — USB Prescriptions with the AFS Pro 1200 Display (Import and Set Up); AFS Pro 1200 Software Operating Manual (Importing Shapefile Data)",
)

add(
    marca="Case IH", monitor="AFS Pro 1200", versao="Todas as versoes",
    objetivo="Importar prescricao (taxa variavel)", origem="Nuvem / plataforma do fabricante",
    formato="Prescricao enviada pelo AFS Connect",
    pasta="Nao se aplica — transferencia sem fio",
    fs="Nao se aplica",
    passos=[
        "Confirme que a maquina esta conectada e visivel no AFS Connect.",
        "No portal AFS Connect, carregue a prescricao e associe ao talhao.",
        "Envie a prescricao para a maquina de destino.",
        "No monitor, aceite/confirme a transferencia recebida.",
        "Abra a configuracao do trabalho e selecione a prescricao.",
        "Defina a coluna de taxa e a unidade.",
        "Confira o mapa antes de iniciar.",
    ],
    cuidados="Depende de cobertura de sinal. Mantenha copia em pen drive.",
    erros="Maquina nao vinculada a conta correta; talhao duplicado na plataforma.",
    confianca="Confirmar",
    fonte="Case IH — How to Load Planting Prescriptions with AFS Connect and the AFS Pro 1200 Display",
)

# ------------------------------------------------------- NEW HOLLAND ------
add(
    marca="New Holland", monitor="IntelliView IV", versao="Todas as versoes",
    objetivo="Importar prescricao (taxa variavel)", origem="Pen drive USB",
    formato="Shapefile (.shp + .shx + .dbf + .prj)",
    pasta="Pasta selecionada na tela de importacao",
    fs="FAT32",
    passos=[
        "Formate o pen drive em FAT32 e copie o shapefile completo para uma pasta de nome simples.",
        "Insira o pen drive na porta USB do IntelliView IV.",
        "Acesse a area de gerenciamento de dados / importacao.",
        "Selecione a importacao de prescricao (Shapefile).",
        "Navegue ate a pasta e selecione o arquivo.",
        "Defina a coluna de taxa e a unidade.",
        "Vincule ao Talhao correto.",
        "Confira o mapa e as taxas antes de iniciar.",
    ],
    cuidados="Mesma plataforma CNH do Case IH AFS Pro 700 — os procedimentos sao praticamente identicos.",
    erros="Shapefile incompleto; pen drive em formato nao suportado.",
    confianca="Confirmar",
    fonte="Plataforma CNH (equivalente ao AFS Pro 700) — confirmar caminho de menu na versao instalada",
)

add(
    marca="New Holland", monitor="IntelliView 12", versao="Todas as versoes",
    objetivo="Importar prescricao (taxa variavel)", origem="Pen drive USB",
    formato="Shapefile (.shp + .shx + .dbf + .prj)",
    pasta="Voce seleciona a PASTA que contem o shapefile na tela de importacao",
    fs="FAT32",
    passos=[
        "Formate o pen drive em FAT32.",
        "Crie uma pasta simples e copie para dentro os 4 arquivos do shapefile.",
        "Insira o pen drive no monitor.",
        "Abra a area de importacao de dados.",
        "Selecione a PASTA que contem o shapefile e confirme.",
        "Escolha a coluna de taxa e a unidade.",
        "Vincule a prescricao ao talhao.",
        "Confira o mapa antes de iniciar.",
    ],
    cuidados="Mesma plataforma CNH do AFS Pro 1200.",
    erros="Selecionar o arquivo em vez da pasta.",
    confianca="Confirmar",
    fonte="Plataforma CNH (equivalente ao AFS Pro 1200)",
)

# ------------------------------------------------------------ TRIMBLE ------
add(
    marca="Trimble", monitor="GFX-750", versao="Todas as versoes",
    objetivo="Importar prescricao (taxa variavel)", origem="Pen drive USB",
    formato="Shapefile (.shp + .shx + .dbf + .prj)",
    pasta="AgData\\Prescriptions\\   (duas pastas: 'AgData' na raiz e 'Prescriptions' dentro dela)",
    fs="FAT32",
    passos=[
        "Formate o pen drive em FAT32.",
        "Na RAIZ do pen drive crie a pasta 'AgData'.",
        "Dentro de 'AgData' crie a pasta 'Prescriptions'.",
        "Copie os 4 arquivos do shapefile para dentro de AgData\\Prescriptions.",
        "Insira o pen drive na porta USB do GFX-750.",
        "No Precision-IQ, abra a tela de transferencia de dados (icone USB / Data Transfer).",
        "Selecione o pen drive como origem e marque a prescricao.",
        "Confirme a importacao.",
        "Ao criar/abrir o trabalho, selecione a prescricao, a coluna de taxa e a unidade.",
        "Confira o mapa e as taxas antes de iniciar.",
    ],
    cuidados="A hierarquia AgData\\Prescriptions e' obrigatoria — fora dela o Precision-IQ nao enxerga o arquivo.",
    erros="Colocar o shapefile solto na raiz; criar so' a pasta 'AgData' sem a subpasta 'Prescriptions'; shapefile incompleto.",
    confianca="Verificado",
    fonte="Trimble Precision-IQ / Farm Data Compatibility; material de suporte Auravant (GFX-350 e GFX-750)",
)

add(
    marca="Trimble", monitor="GFX-1060 / GFX-1260", versao="Todas as versoes",
    objetivo="Importar prescricao (taxa variavel)", origem="Pen drive USB",
    formato="Shapefile (.shp + .shx + .dbf + .prj)",
    pasta="AgData\\Prescriptions\\",
    fs="FAT32",
    passos=[
        "Formate o pen drive em FAT32.",
        "Crie a pasta 'AgData' na raiz e dentro dela a pasta 'Prescriptions'.",
        "Copie os 4 arquivos do shapefile para AgData\\Prescriptions.",
        "Insira o pen drive no monitor.",
        "No Precision-IQ, abra a tela de transferencia de dados.",
        "Selecione o pen drive e a prescricao; confirme a importacao.",
        "No trabalho, selecione a prescricao, a coluna de taxa e a unidade.",
        "Confira o mapa antes de iniciar.",
    ],
    cuidados="Mesma logica de pastas do GFX-750.",
    erros="Hierarquia de pastas incorreta.",
    confianca="Verificado",
    fonte="Trimble Precision-IQ — mesma estrutura AgData do GFX-750",
)

add(
    marca="Trimble", monitor="TMX-2050", versao="Todas as versoes",
    objetivo="Importar prescricao (taxa variavel)", origem="Pen drive USB",
    formato="ISOXML (TASKDATA.XML + arquivos .BIN)",
    pasta="TASKDATA\\  (na RAIZ do pen drive)",
    fs="FAT32",
    passos=[
        "No escritorio, exporte o trabalho no formato ISOXML (gera um .zip).",
        "Descompacte o .zip: dentro ha o TASKDATA.XML e possiveis arquivos .BIN.",
        "Formate o pen drive em FAT32.",
        "Crie a pasta 'TASKDATA' na RAIZ do pen drive e copie para dentro o TASKDATA.XML e os .BIN.",
        "Insira o pen drive no monitor.",
        "Abra a tela de transferencia/importacao de dados.",
        "Selecione o pen drive e importe a tarefa (task).",
        "Abra a tarefa importada e confira o mapa de taxas antes de iniciar.",
    ],
    cuidados="Nao renomeie TASKDATA.XML nem os arquivos .BIN. A pasta deve se chamar exatamente TASKDATA.",
    erros="Deixar o .zip sem descompactar; criar TASKDATA dentro de outra pasta; renomear arquivos.",
    confianca="Verificado",
    fonte="Trimble — TASKDATA (ISOXML) no TMX-2050; documentacao ISOBUS",
)

add(
    marca="Trimble", monitor="GFX-750", versao="Todas as versoes",
    objetivo="Exportar dados de trabalho (as-applied / colheita)", origem="Pen drive USB",
    formato="Pacote AgData (.zip) gerado pelo Precision-IQ",
    pasta="AgData\\  (o proprio monitor cria a estrutura)",
    fs="FAT32",
    passos=[
        "Finalize o trabalho no Precision-IQ.",
        "Insira o pen drive na porta USB.",
        "Abra a tela de transferencia de dados (Data Transfer).",
        "Selecione os trabalhos/talhoes que deseja exportar.",
        "Escolha o pen drive como destino e confirme a exportacao.",
        "Aguarde a conclusao antes de remover o pen drive.",
        "No PC, importe a pasta AgData no Trimble Ag Software ou no SMS/software de gestao.",
    ],
    cuidados="Mantenha a pasta AgData inteira ao copiar para o PC.",
    erros="Copiar apenas parte da estrutura; remover o pen drive antes do fim.",
    confianca="Confirmar",
    fonte="Trimble Precision-IQ — confirmar nomes de tela na versao instalada",
)

# ---------------------------------------------------------- AG LEADER ------
add(
    marca="Ag Leader", monitor="InCommand 800 / InCommand 1200", versao="Todas as versoes",
    objetivo="Importar limites e cadastro (cliente/fazenda/talhao)", origem="Pen drive USB",
    formato=".agsetup  (contem cliente/fazenda/talhao, limites, linhas de guia e prescricoes)",
    pasta="Raiz do pen drive (ou qualquer pasta — voce localiza o arquivo na tela)",
    fs="FAT32",
    passos=[
        "No SMS/AgFiniti, exporte os dados de configuracao como arquivo .agsetup.",
        "Copie o .agsetup para o pen drive.",
        "Insira o pen drive na porta USB do InCommand.",
        "Abra a pagina 'Transferencia de Dados' (Data Transfer).",
        "Toque em 'Importar Setup' (Import Setup).",
        "Localize o arquivo .agsetup no pen drive e selecione.",
        "Siga as instrucoes e escolha o que deseja importar (talhoes, limites, linhas, prescricoes).",
        "Confirme e aguarde a conclusao.",
        "Verifique em Configuracao se os talhoes e linhas apareceram corretamente.",
    ],
    cuidados="O .agsetup so' e' trocado entre grupos de display compativeis: Integra/Versa entre si e InCommand entre si.",
    erros="Tentar importar .agsetup de Integra/Versa direto no InCommand; para mover PADROES entre geracoes use o arquivo .pat.",
    confianca="Verificado",
    fonte="Ag Leader — InCommand User Guide (Import .agsetup) e portal Ag Leader 'AgSetup File Supported Uses'",
)

add(
    marca="Ag Leader", monitor="InCommand 800 / InCommand 1200", versao="Todas as versoes",
    objetivo="Importar linhas de guia (AB / curvas)", origem="Pen drive USB",
    formato=".agsetup (entre displays do mesmo grupo)  |  .pat (para mover padroes entre geracoes)",
    pasta="Raiz do pen drive",
    fs="FAT32",
    passos=[
        "Defina a origem: se for entre displays InCommand, use .agsetup; se for de Integra/Versa para InCommand, use .pat.",
        "Exporte o arquivo no display/software de origem.",
        "Copie o arquivo para o pen drive.",
        "Insira o pen drive no InCommand.",
        "Abra 'Transferencia de Dados' > 'Importar Setup'.",
        "Selecione o arquivo e escolha importar as linhas de guia (padroes).",
        "Confirme e aguarde.",
        "Abra o talhao e selecione a linha de guia importada na tela de Guia.",
        "Confira a largura de trabalho e o alinhamento em campo antes de iniciar.",
    ],
    cuidados="Esta e' a pegadinha classica da Ag Leader: geracao diferente exige .pat, nao .agsetup.",
    erros="Usar .agsetup entre geracoes diferentes e concluir que 'o arquivo esta corrompido'.",
    confianca="Verificado",
    fonte="Ag Leader — 'AgSetup File Supported Uses'; Precision Consulting: Transferring Guidance Lines Between Displays",
)

add(
    marca="Ag Leader", monitor="InCommand 800 / InCommand 1200", versao="Todas as versoes",
    objetivo="Exportar dados de trabalho (as-applied / colheita)", origem="Pen drive USB",
    formato=".agdata  (dados registrados)",
    pasta="Raiz do pen drive",
    fs="FAT32",
    passos=[
        "Encerre o trabalho no display.",
        "Insira o pen drive na porta USB.",
        "Abra a pagina 'Transferencia de Dados' (Data Transfer).",
        "Selecione 'Exportar' e escolha os dados registrados.",
        "O display gera um arquivo .agdata no pen drive.",
        "Aguarde a conclusao antes de remover o pen drive.",
        "No PC, importe o .agdata no SMS ou envie para o AgFiniti.",
    ],
    cuidados="Exporte com frequencia. Nao dependa apenas da memoria interna do display.",
    erros="Confundir .agsetup (configuracao) com .agdata (dados registrados).",
    confianca="Verificado",
    fonte="Ag Leader — InCommand 800/1200 User Guide (Exporting Logged Data / AgData files)",
)

# ------------------------------------------------------------- TOPCON ------
add(
    marca="Topcon", monitor="X35 / XD+ (Horizon)", versao="Todas as versoes",
    objetivo="Importar prescricao (taxa variavel)", origem="Pen drive USB",
    formato="ISOXML (TASKDATA.XML + arquivos .BIN)",
    pasta="TASKDATA\\  (na RAIZ do pen drive)",
    fs="FAT32 (obrigatorio)",
    passos=[
        "No software de escritorio, exporte a tarefa no formato ISOXML (gera um .zip).",
        "Descompacte o .zip — dentro ha o TASKDATA.XML e possiveis arquivos .BIN.",
        "Formate o pen drive em FAT32.",
        "Crie a pasta 'TASKDATA' na RAIZ e copie para dentro o TASKDATA.XML e os .BIN.",
        "Configure o terminal para operar com ISOXML (implemento ISOBUS reconhecido no Universal Terminal).",
        "Insira o pen drive na porta USB do terminal.",
        "Abra o gerenciador de tarefas (Task Controller) e importe a tarefa do pen drive.",
        "Selecione a tarefa importada e confira o mapa de taxas.",
        "Verifique se o implemento esta em TC-GEO (taxa variavel por posicao) antes de iniciar.",
    ],
    cuidados="Sem TC-GEO habilitado o terminal aceita a tarefa mas nao aplica taxa variavel.",
    erros="Pen drive em exFAT; TASKDATA dentro de outra pasta; arquivos .BIN esquecidos.",
    confianca="Verificado",
    fonte="Topcon X35 Operator's Manual (TASKDATA via USB, FAT32); documentacao ISOBUS TC-BAS/TC-GEO",
)

add(
    marca="AGCO (Fendt / Valtra / Massey Ferguson)",
    monitor="Varioterminal / FendtONE / terminal ISOBUS AGCO", versao="Todas as versoes",
    objetivo="Importar prescricao (taxa variavel)", origem="Pen drive USB",
    formato="ISOXML (TASKDATA.XML + arquivos .BIN)",
    pasta="TASKDATA\\  (na RAIZ do pen drive)",
    fs="FAT32",
    passos=[
        "No software de escritorio, exporte o mapa de aplicacao em ISOXML (arquivo .zip).",
        "Descompacte o .zip: TASKDATA.XML + arquivos .BIN.",
        "Formate o pen drive em FAT32.",
        "Crie a pasta 'TASKDATA' na RAIZ e copie os arquivos para dentro.",
        "Insira o pen drive na porta USB do terminal.",
        "Abra a area de tarefas / TaskDoc no terminal.",
        "Importe a tarefa do pen drive.",
        "Selecione a tarefa e confira o mapa de taxas.",
        "Confirme que o implemento ISOBUS esta em TC-GEO antes de iniciar.",
    ],
    cuidados="Terminais AGCO seguem ISOBUS puro (TC-BAS e TC-GEO). O nome do menu muda entre Varioterminal e FendtONE.",
    erros="Deixar o .zip sem descompactar; pasta TASKDATA aninhada; implemento sem licenca de taxa variavel.",
    confianca="Verificado",
    fonte="AGCO / Fendt — carregamento de mapas de taxa variavel via ISOXML; documentacao TaskDoc / VarioDoc",
)

add(
    marca="AGCO (Fendt / Valtra / Massey Ferguson)",
    monitor="Varioterminal / FendtONE / terminal ISOBUS AGCO", versao="Todas as versoes",
    objetivo="Exportar dados de trabalho (as-applied / colheita)", origem="Pen drive USB",
    formato="ISOXML (TASKDATA.XML atualizado com os registros + .BIN)",
    pasta="TASKDATA\\  (o terminal grava na mesma estrutura)",
    fs="FAT32",
    passos=[
        "Encerre a tarefa no terminal (status 'concluida').",
        "Insira o pen drive na porta USB.",
        "Abra a area de tarefas / TaskDoc.",
        "Selecione exportar/gravar tarefas no pen drive.",
        "O terminal atualiza a pasta TASKDATA com os dados registrados.",
        "Aguarde a conclusao antes de remover o pen drive.",
        "No PC, compacte a pasta TASKDATA em .zip e importe no software de gestao.",
    ],
    cuidados="Nao encerrar a tarefa antes de exportar e' a causa mais comum de dado faltando.",
    erros="Exportar com a tarefa aberta; copiar so' o TASKDATA.XML sem os .BIN.",
    confianca="Confirmar",
    fonte="AGCO TaskDoc / VarioDoc — confirmar nomes de menu na linha e versao do terminal",
)

# -------------------------------------------------------------- RAVEN ------
add(
    marca="Raven", monitor="Viper 4 / Viper 4+", versao="Todas as versoes",
    objetivo="Importar prescricao (taxa variavel)", origem="Pen drive USB",
    formato="Shapefile (.shp + .shx + .dbf + .prj)",
    pasta="Estrutura gerada/reconhecida pelo ROS — confirme na versao instalada",
    fs="FAT32",
    passos=[
        "Formate o pen drive em FAT32.",
        "Copie o shapefile completo para o pen drive.",
        "Insira o pen drive na porta USB do Viper 4.",
        "Acesse a area de gerenciamento de arquivos do ROS.",
        "Selecione importar prescricao e localize o arquivo.",
        "Escolha a coluna de taxa e a unidade.",
        "Vincule ao talhao/trabalho.",
        "Confira o mapa e as taxas antes de iniciar.",
    ],
    cuidados="A estrutura de pastas do Viper 4 varia entre versoes do ROS. Registre o caminho real da sua maquina.",
    erros="Shapefile incompleto; projecao diferente de WGS84.",
    confianca="Confirmar",
    fonte="Raven ROS — confirmar estrutura e caminho de menu na versao instalada",
)


PROCEDIMENTOS = P

# =========================================================================== #
#  5. FORMATOS DE ARQUIVO                                                     #
# =========================================================================== #
FORMATOS = [
    # (formato, extensoes, o que carrega, quem usa, cuidados)
    ("Shapefile (ESRI)", ".shp + .shx + .dbf + .prj (+ .cpg opcional)",
     "Prescricoes de taxa variavel, limites de talhao, poligonos de zona",
     "John Deere (pasta Rx), Case IH, New Holland, Trimble (AgData\\Prescriptions), Raven, Ag Leader",
     "Os 4 arquivos devem ter o MESMO nome-base e viajar juntos. O .prj carrega a projecao — sem ele o monitor pode nao posicionar o mapa. Use WGS84 e geometria de POLIGONO. A coluna de taxa precisa ser NUMERICA."),

    ("ISOXML (ISO 11783-10)", "TASKDATA.XML + arquivos .BIN (dentro da pasta TASKDATA)",
     "Tarefas completas: talhao, limites, prescricao, produto, registros da operacao",
     "Topcon, AGCO (Fendt/Valtra/Massey), CLAAS, Trimble TMX-2050, varios ISOBUS",
     "Padrao aberto ISOBUS. A pasta deve se chamar exatamente TASKDATA e ficar na RAIZ. Nunca renomeie os arquivos. Exige TC-GEO no implemento para taxa variavel."),

    (".agsetup (Ag Leader)", ".agsetup",
     "Configuracao: cliente/fazenda/talhao, limites, LINHAS DE GUIA e prescricoes",
     "Ag Leader (Integra, Versa, InCommand) e SMS/AgFiniti",
     "So' e' trocado entre grupos de display compativeis. Integra/Versa entre si; InCommand entre si."),

    (".agdata (Ag Leader)", ".agdata",
     "Dados REGISTRADOS da operacao (as-applied, colheita)",
     "Ag Leader InCommand / Integra / Versa e SMS",
     "Nao confunda com .agsetup. O .agdata e' o resultado do trabalho, nao a configuracao."),

    (".pat (Ag Leader)", ".pat",
     "Padroes de guia (linhas AB, curvas) isolados",
     "Ag Leader — usado para mover padroes ENTRE GERACOES de display",
     "E' o formato correto quando o .agsetup nao e' aceito (ex.: Integra/Versa -> InCommand)."),

    ("Pacote AgData (Trimble)", "pasta AgData\\ (as vezes distribuida como .zip)",
     "Trabalhos, talhoes, linhas de guia e prescricoes do Precision-IQ",
     "Trimble GFX-750, GFX-1060/1260, TMX-2050 (Precision-IQ)",
     "Prescricoes shapefile ficam em AgData\\Prescriptions. Mantenha a estrutura inteira ao copiar."),

    ("Pacote AgGPS (Trimble legado)", "pasta AgGPS\\",
     "Dados dos monitores FmX / FmX+ (geracao anterior ao Precision-IQ)",
     "Trimble FmX, FmX+",
     "Geracao antiga. Nao e' intercambiavel com AgData sem conversao."),

    ("Estrutura GreenStar 3", "GS3_2630\\<Perfil>\\RCD\\",
     "Dados gravados e perfis do GreenStar 3 2630",
     "John Deere GreenStar 3 2630",
     "Copie a pasta GS3_2630 INTEIRA. Copiar so' o conteudo do RCD quebra a leitura no escritorio."),

    ("Estrutura Gen 4 / G5", "Rx\\ (entrada de prescricao)  |  JD-Data\\ (dados de trabalho)  |  JD4600\\ (configuracao)",
     "Prescricoes de entrada e dados exportados dos displays Gen 4 e G5",
     "John Deere Gen 4 (4240/4600/4640) e G5",
     "A pasta Rx tem que estar na RAIZ do pen drive. As pastas JD-Data e JD4600 sao criadas pelo proprio monitor na exportacao."),

    ("GeoTIFF / imagem raster", ".tif, .tiff (+ .tfw)",
     "Imagens de fundo, mapas de produtividade em raster, indices (NDVI)",
     "Softwares de escritorio; poucos monitores importam diretamente",
     "Normalmente precisa ser convertido em poligonos (shapefile) antes de virar prescricao."),

    ("CSV / TXT", ".csv, .txt",
     "Pontos de amostragem, tabelas de taxa, dados de estacao meteorologica",
     "Softwares de escritorio e esta planilha (dados da estacao)",
     "Confirme o separador decimal (virgula x ponto) e o separador de campos antes de importar."),
]
