# NTS Rotas

Plataforma web da NTS para importar relatórios do Spoke Dispatch, calcular a quilometragem válida e o bônus dos motoboys e exportar o fechamento do período.

Todo o processamento acontece diretamente no navegador. O projeto não possui backend, banco de dados, autenticação, SSR ou envio do Excel para servidores.

## Tecnologias utilizadas

- React
- TypeScript
- Vite
- Tailwind CSS
- Recharts
- SheetJS (`xlsx`)
- Vitest

## Como executar localmente

Requisito: Node.js 22 ou superior.

```bash
pnpm install
pnpm dev
```

Abra o endereço exibido pelo Vite e selecione um arquivo `.xlsx` ou `.xls`.

Para executar os testes e validar o build:

```bash
pnpm test
pnpm run build
pnpm preview
```

O build de produção é gerado na pasta `dist`.

## Como publicar na Vercel

1. Importe o repositório na Vercel.
2. Selecione Vite como framework.
3. Use `pnpm run build` como comando de build.
4. Use `dist` como diretório de saída.
5. Não configure variáveis de ambiente.

O arquivo `vercel.json` direciona rotas da SPA para `index.html`.

## Estrutura de pastas

```text
src/
  App.tsx          Interface e interações
  index.css        Identidade visual e responsividade
  main.tsx         Entrada da aplicação React
lib/
  closing.ts       Leitura, tratamento, cálculos e exportação
  closing.test.ts  Testes das regras principais
  reference-data.test.ts  Homologação com o Excel de referência
public/            Favicon e arquivos públicos
dist/              Build estático gerado pelo Vite
```

## Regras de negócio

- A distância é lida de `distance_km`.
- No export atual do Spoke, somente `stop_state = delivered` é considerado concluído.
- Quando o arquivo utiliza o schema alternativo, somente `status = completed` é considerado concluído.
- Outros estados são excluídos.
- Distâncias vazias, inválidas ou negativas são excluídas e auditadas.
- Entregas são identificadas por `tracking_code`.
- Trechos de saída e retorno à base sem `tracking_code` entram na quilometragem, mas não aumentam o total de entregas.
- Registros duplicados são desconsiderados.
- Nomes de motoboys são normalizados para evitar duplicação por espaços, caixa ou acentuação equivalente.
- O bônus é a quilometragem válida multiplicada pelo valor vigente por quilômetro.
- O valor inicial é R$ 0,25 por quilômetro e fica salvo no `localStorage`.

## Fluxo do usuário

1. Acessar a plataforma.
2. Selecionar ou arrastar o Excel exportado do Spoke.
3. Aguardar o processamento local.
4. Conferir indicadores, ranking, evolução diária e auditoria.
5. Pesquisar, ordenar ou abrir o detalhamento de um motoboy.
6. Ajustar o valor por quilômetro, se necessário.
7. Exportar o fechamento em Excel.

## Versão atual

Versão 1.0.1 — polimento da interface e documentação da versão aprovada.

Consulte o [CHANGELOG.md](./CHANGELOG.md) para o histórico publicado.
