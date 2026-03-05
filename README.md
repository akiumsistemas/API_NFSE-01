# 📥 Baixador de XMLs (DF-e) - Portal Nacional

Este projeto é um script em Node.js automatizado para conectar à API de Distribuição de Documentos Fiscais Eletrônicos (DF-e) do Governo (Ambiente de Produção).

Ele realiza a autenticação via **Certificado Digital A1**, baixa o lote de documentos, descompacta o conteúdo (GZIP/Base64) e salva os arquivos `.xml` localmente.

## 🚀 Funcionalidades

* 🔐 Autenticação SSL mútua com certificado `.pfx`.
* 📡 Conexão direta com a URL de Produção do Governo.
* 📦 Descodificação automática de Base64.
* 🗜️ Descompactação automática de GZIP (padrão do Portal Nacional).
* 💾 Salvamento automático do XML com o nome da Chave de Acesso.

## 📋 Pré-requisitos

* [Node.js](https://nodejs.org/) instalado (v14 ou superior).
* Arquivo do Certificado Digital (**modelo A1**, extensão `.pfx` ou `.p12`).
* Senha do certificado.

## 🛠️ Instalação

1. Clone este repositório ou baixe os arquivos.
2. Na pasta do projeto, instale as dependências:

```bash
npm install axios dotenv

## Comando para executar ##

node --openssl-legacy-provider index.js