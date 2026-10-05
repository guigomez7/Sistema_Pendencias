# Sistema de Envio Automático de Pendências

Sistema 100% local para Windows (sem site, sem banco de dados, sem hospedagem).
Lê os arquivos de pendências em Excel, identifica a empresa e a data pelo nome
do arquivo, monta o e-mail correspondente com o Excel anexado e envia via SMTP.

## 1. Estrutura de pastas

Copie a pasta **`sistema_pendencias`** inteira para a **Área de Trabalho** do
Windows. A estrutura final deve ficar assim:

```
Área de Trabalho/
└── sistema_pendencias/
    ├── enviar_pendencias.py
    ├── enviar_pendencias.bat   <- é este que você clica duas vezes
    ├── config.json
    ├── credenciais.json         (criado automaticamente no primeiro uso)
    ├── requirements.txt
    ├── pendencias/               <- coloque aqui os arquivos "Pendências [Empresa] [Data].xlsx"
    ├── empresas/
    │   └── Empresas.xlsx         <- coluna A = nome da empresa, coluna B = e-mail
    ├── enviados/                 <- criado automaticamente (histórico por data)
    ├── erros/                    <- criado automaticamente
    └── logs/                     <- criado automaticamente
```

## 2. Configurar o e-mail de envio (config.json)

Abra `config.json` em um editor de texto e ajuste:

- `smtp.host`: endereço do servidor SMTP da CROCK.cc (confirme com o painel/
  suporte da hospedagem de e-mail; o valor padrão é um palpite razoável).
- `smtp.porta`: normalmente `587` (com TLS) ou `465` (SSL). O script detecta
  automaticamente o método correto pela porta informada.
- `smtp.remetente_email`: o e-mail completo que vai enviar as mensagens.
- `smtp.remetente_nome`: nome que aparece como remetente.
- `modo_teste_padrao`: `true` para o sistema sugerir modo teste por padrão.

**Nunca coloque a senha neste arquivo.**

## 3. Senha do e-mail (credenciais.json)

A senha **não fica no código nem no config.json**. Na primeira execução em
modo de envio real, o sistema pede a senha no terminal (ela não aparece
digitada na tela) e a salva em `credenciais.json`. Mantenha esse arquivo
apenas neste computador — não o envie por e-mail, não o suba para nuvens
públicas nem para repositórios (ex.: GitHub).

## 4. Cadastro de empresas (empresas/Empresas.xlsx)

Já está incluída uma planilha inicial com as empresas e e-mails cadastrados.
Basta manter essa planilha atualizada:

- Coluna A: nome da empresa (deve bater com o nome usado no nome do arquivo
  de pendências, sem se preocupar com acentos/maiúsculas — a comparação
  ignora isso).
- Coluna B: e-mail de destino.

## 5. Padrão do nome dos arquivos de pendências

```
Pendências [Nome da Empresa] [DDMMAAAA].xlsx
```

Exemplo: `Pendências AcessoCar 10092026.xlsx`

Arquivos fora desse padrão (data com número errado de dígitos, data
inválida, etc.) são automaticamente separados como erro e não são enviados.

## 6. Como usar no dia a dia

1. Coloque os arquivos `.xlsx` de pendências na pasta `pendencias`.
2. Dê **dois cliques** em `enviar_pendencias.bat`.
3. O sistema vai:
   - Instalar automaticamente a biblioteca necessária (`openpyxl`), se faltar;
   - Mostrar um **resumo** de tudo que encontrou: empresas identificadas,
     e-mails, arquivos com problema e duplicados;
   - Perguntar se você quer rodar em **modo teste** (nada é enviado de
     verdade) ou enviar direto;
   - Pedir **confirmação** antes de enviar qualquer e-mail;
   - Enviar os e-mails um a um, mostrando o progresso no terminal — se um
     falhar, os demais continuam normalmente;
   - Mover os arquivos enviados com sucesso para `enviados/[data]` e os
     arquivos com problema para `erros`;
   - Registrar tudo em `logs` (histórico completo e um log detalhado da
     execução), evitando reenviar a mesma pendência da mesma empresa/data.

## 7. Modo teste

Use o modo teste sempre que quiser conferir se os arquivos estão sendo
identificados corretamente, sem risco de disparar e-mails reais. Você pode
forçá-lo diretamente pela linha de comando:

```
python enviar_pendencias.py --teste
python enviar_pendencias.py --enviar
```

(o `.bat` também aceita a pergunta interativa normalmente, sem precisar
digitar comandos).

## 8. Segurança

- A senha nunca é exibida no terminal nem gravada no código-fonte.
- `credenciais.json` fica separado de `config.json` justamente para poder
  ser excluído/protegido sem mexer nas demais configurações.
- Nenhum dado é enviado para fora do seu computador, exceto o próprio
  e-mail (via SMTP) para o destinatário cadastrado.
