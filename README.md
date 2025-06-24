# Análise de Curvas de Bomba

Um aplicativo Python para análise e processamento de curvas características de bombas centrífugas, com interface gráfica moderna e geração automática de relatórios Excel.

## 📋 Funcionalidades

- **Importação de dados** a partir de imagens de gráficos
- **Entrada manual de dados** com interface intuitiva
- **Análise de múltiplos rotores** com diferentes diâmetros
- **Cálculo de curvas do sistema** (manual ou por equação)
- **Geração automática de relatórios Excel** com gráficos
- **Cálculo de pontos de interseção** entre curvas de bomba e sistema
- **Suporte a bombas em paralelo**
- **Alteração de RPM** com aplicação das leis de afinidade

## 🚀 Tecnologias Utilizadas

- **Python 3.x**
- **PyQt5** - Interface gráfica
- **OpenPyXL** - Manipulação de arquivos Excel
- **NumPy** - Cálculos numéricos
- **SciPy** - Interpolação e otimização
- **Matplotlib** - Geração de gráficos

## 📦 Instalação

1. Clone o repositório:
```bash
git clone https://github.com/skinerbold/Graficos_Bomba_d-agua.git
cd Graficos_Bomba_d-agua
```

2. Instale as dependências:
```bash
pip install PyQt5 openpyxl numpy scipy matplotlib
```

3. Execute o aplicativo:
```bash
python main.py
```

## 💡 Como Usar

### Modo de Entrada Manual
1. Selecione "Entrada Manual de Dados" na tela inicial
2. Adicione rotores usando o botão "Adicionar Rotor"
3. Insira os dados de vazão, altura e eficiência para cada rotor
4. Configure as curvas do sistema (opcional)
5. Gere o relatório Excel

### Modo de Importação de Imagem
1. Selecione "Importar de Imagem" na tela inicial
2. Carregue uma imagem com gráficos de curvas de bomba
3. Use as ferramentas de calibração para definir os eixos
4. Marque os pontos das curvas
5. Gere o relatório Excel

## 📊 Recursos do Relatório

O relatório Excel gerado inclui:
- Planilha com dados originais
- Planilha com dados interpolados
- Planilha com pontos de interseção
- Gráficos das curvas de bomba e sistema
- Análise de eficiência máxima

## 🔧 Funcionalidades Avançadas

- **Bombas em Paralelo**: Crie automaticamente curvas para bombas operando em paralelo
- **Alteração de RPM**: Aplique as leis de afinidade para diferentes rotações
- **Múltiplas Curvas de Sistema**: Compare até duas curvas de sistema diferentes
- **Cálculo Automático de Interseções**: Encontre pontos de operação automaticamente

## 📝 Estrutura do Projeto

```
curvas-bomba/
├── main.py              # Arquivo principal
├── requirements.txt     # Dependências
├── README.md           # Este arquivo
├── .gitignore          # Arquivos ignorados pelo Git
└── exemplos/           # Arquivos de exemplo (se houver)
```

## 🤝 Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para:
- Reportar bugs
- Sugerir novas funcionalidades
- Enviar pull requests

## 📄 Licença

Este projeto está sob a licença MIT. Veja o arquivo `LICENSE` para mais detalhes.

## ✨ Autor

Desenvolvido por Skiner Bold

---

⭐ Se este projeto foi útil para você, considere dar uma estrela no repositório!
