# Gerador de GeoJSON para ARTESP

Ferramenta desenvolvida para automação de relatórios técnicos de conservação e obras em rodovias.

## 🛠️ Tecnologias
- Python 3.10
- GeoPandas (Manipulação de dados espaciais)
- SignTool (Assinatura Digital)

## 🚀 Como usar
1. Instale as dependências: `conda env update --file environment.yml`.
2. (Opcional, recomendado para equipe) Configure assinatura digital com 1 clique:
   - execute `configurar_artesp_1clique.bat` (duplo clique);
   - o assistente cria `C:\ARTESP`, copia o `.pfx` e grava `ARTESP_PFX`/`ARTESP_PFX_PASSWORD`.
   - alternativa para CI/Render: use `ARTESP_PFX_CONTENT` (Base64 do `.pfx`) + `ARTESP_PFX_PASSWORD`.
   - quando `ARTESP_PFX_CONTENT` for usado, o build reconstrói um `.pfx` temporário, assina e remove o arquivo temporário.
   - para automacao de TI (sem prompts), use modo silencioso no PowerShell:
     - `.\setup_artesp_env.ps1 -Silent -PfxSourcePath "C:\Deploy\certificado.pfx" -PfxPassword (ConvertTo-SecureString "SUA_SENHA" -AsPlainText -Force) -NoPause`
     - ou defina `ARTESP_PFX_PASSWORD_BOOTSTRAP` e rode:
       `.\setup_artesp_env.ps1 -Silent -PfxSourcePath "C:\Deploy\certificado.pfx" -NoPause`
   - para distribuicao via GPO/Intune, use:
     - `.\deploy_artesp_silent.ps1 -PfxSourcePath "C:\Deploy\certificado.pfx" -BootstrapPassword (ConvertTo-SecureString "SUA_SENHA" -AsPlainText -Force)`
     - logs: `C:\ARTESP\outputs\deploy_logs`
     - exit codes: `0` sucesso, `11` sem admin, `12` setup ausente, `13` pfx ausente, `14` pfx invalido, `15` senha ausente, `31` falha setup, `99` erro inesperado.
   - para testar pipeline sem aplicar alteracoes, rode dry-run:
     - `.\deploy_artesp_silent.ps1 -PfxSourcePath "C:\Deploy\certificado.pfx" -BootstrapPassword (ConvertTo-SecureString "SUA_SENHA" -AsPlainText -Force) -DryRun`
     - no dry-run, o script apenas valida prerequisitos e nao grava variaveis.
3. Gere o executável com `python build.py` ou `python build_exe.py`.
4. Se as variáveis estiverem configuradas, a assinatura é automática:
   - Windows: `signtool`
   - Linux (Docker/Render): `osslsigncode`
5. A cada assinatura concluída, é criado um `.log` em `outputs/` com o `ThumbprintSHA1` do certificado usado.
6. Para diagnóstico rápido de ambiente, execute `python check_structure.py`.
7. Para gerar `ARTESP_PFX_CONTENT` (Base64 do `.pfx`), use:
   - utilitario dedicado: `python utils/export_pfx_base64.py "C:\ARTESP\certificado.pfx" --out "outputs\ARTESP_PFX_CONTENT.txt"`
   - via função administrativa do diagnóstico:
     `python check_structure.py --export-pfx-base64 "C:\ARTESP\certificado.pfx" --out "outputs\ARTESP_PFX_CONTENT.txt"`
   - atalho para equipe (duplo clique): `exportar_pfx_base64_equipe.bat` (usa `ARTESP_PFX` automaticamente).

## 👷 Autor
Desenvolvido por um especialista em conservação rodoviária com foco em automação e eficiência de dados.
