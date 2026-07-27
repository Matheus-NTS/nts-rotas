# NTS Rotas

Plataforma web para importar o Excel do Spoke Dispatch, calcular a quilometragem válida e o bônus dos motoboys e exportar o fechamento.

## Executar localmente

Requisitos: Node.js 22 ou superior.

```bash
pnpm install
pnpm dev
```

Abra o endereço exibido no terminal, selecione um arquivo `.xlsx` ou `.xls` e confira o fechamento.

## Testes e build

```bash
pnpm test
pnpm build
pnpm preview
```

## Regras implementadas

- Processamento integralmente no navegador, sem backend ou banco de dados.
- Usa `distance_km` por trecho da rota. Trechos de saída e retorno à base são mantidos; não são totais duplicados.
- Usa `status = completed` quando essa coluna existe.
- O export atual de referência não possui `status`; nele, a conclusão é representada exatamente por `stop_state = delivered`. Esse schema é detectado automaticamente.
- Distâncias vazias, inválidas e negativas são excluídas e auditadas.
- Entregas são identificadas por `tracking_code`; os trechos de base não possuem esse código e não aumentam o total de entregas.
- Registros repetidos são excluídos por `tracking_code` ou, nos trechos de base, pela identidade composta de motoboy, rota, data, parada, estado, distância e endereço.
- Nomes de motoboys são normalizados para evitar duplicação por espaços, caixa ou acentuação simples.
- O valor por quilômetro começa em R$ 0,25, é editável e fica salvo no navegador.
- Exportação com quatro abas: resumo, detalhamento diário, registros considerados e desconsiderados.

## Publicar na Vercel

Importe o repositório na Vercel, use o comando de build `pnpm build` e o diretório de saída `dist`. A aplicação é uma SPA estática criada com React e Vite, sem SSR, backend ou variáveis de ambiente.
