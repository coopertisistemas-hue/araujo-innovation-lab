# DELTA_REPORT — candidato d36fefb → candidato de fechamento (ordem §2.1/§2.2)

| Item | d36fefb (base) | Candidato de fechamento |
|---|---|---|
| Markdown V5.1 | `13d5bc05d0026486f230076901e1dedbda83f6cff52e452c485ba8fd915dda23` | `48f5b2929583e73277759d68314a43528cf0e9961bd551f3e0a761815e6f052e` |
| DOCX | arquivado em `output/historico-v5.1-d36fefb/` | `46d2ca8d00d446505ed446cfad1f18729e341759c753791f35f76545347b9853` |
| PDF | arquivado em `output/historico-v5.1-d36fefb/` | `fa4a1cddd497b24afda7fa2a39b967b02e5d81f08cb95e6e63e121e79b28fbcf` (101 páginas) |
| Testes | 32/32 | 32/32 (com 101 páginas esperadas) |

## Alteração editorial
- §2.1 (Mapa integrado dos pilares, BS10) passa a vir antes de §2.2 (Objetivos e perfil profissional oficiais, PPC II–IV).
- Nenhum conteúdo foi alterado: o conjunto de linhas do Markdown é idêntico ao da base (verificado por comparação multiconjunto de linhas).
- Gerador (`tools/apply_v5_1.py`) corrigido para inserir §2.2 antes de §3; reconstruções futuras mantêm a ordem.

## Preservação
- 183 itens do PPC, 49 unidades de aprendizagem, 28 requisitos do brainstorming e as duas coortes de 2027: sem alteração de conteúdo.
- Cargas horárias e ementas oficiais: sem alteração.

## Efeitos colaterais de layout
- Paginação passou de 100 para 101 páginas; o sumário foi regenerado em duas passagens e está estável.
- Páginas com pouco texto: 19, 52, 66 — fins de seção, intencionais.
