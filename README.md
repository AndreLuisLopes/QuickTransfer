# QuickTransfer

## Descrição
O **QuickTransfer** é uma aplicação desenvolvida em **Python** para facilitar a transferência de dados entre planilhas Excel (`.xlsx`).  
Ele permite selecionar arquivos de origem e destino, aplicar filtros, escolher colunas específicas para copiar e até separar valores com base em um caractere especial, direcionando cada parte para colunas distintas.

---

## Funcionalidades
- Seleção de planilhas de origem e destino.  
- Visualização prévia dos dados antes da transferência.  
- Filtro opcional para restringir os valores transferidos.  
- Suporte à separação de valores com base em um caractere delimitador.  
- Interface gráfica simples e intuitiva utilizando **Tkinter**.  
- Manipulação e leitura de arquivos Excel com **Pandas** e **OpenPyXL**.  

---

## Tecnologias Utilizadas
- **Python 3.x**  
- **Tkinter** (interface gráfica)  
- **Pandas** (manipulação de dados)  
- **OpenPyXL** (leitura e escrita em arquivos Excel)  

---

## Instalação

Clone este repositório:

```bash
git clone https://github.com/seu-usuario/quicktransfer.git
```

Acesse a pasta do projeto:

```bash
cd quicktransfer
```

Instale as dependências necessárias:

```bash
pip install -r requirements.txt
```

---

## Como Usar

Execute o programa:

```bash
python quicktransfer.py
```

Passos básicos:
1. Selecione o arquivo de origem (`.xlsx`).  
2. Selecione o arquivo de destino (`.xlsx`) e a aba correspondente.  
3. Aplique filtros (opcional).  
4. Configure o caractere separador para dividir valores em duas colunas (opcional).  
5. Visualize a prévia e confirme a transferência.  

---

## Autor
Desenvolvido por André Luís Lopes
[GitHub](https://github.com/AndreLuisLopes) • [Linkedin](https://www.linkedin.com/in/andre-luis-lopes/)  
Licenciado sob [MIT](LICENSE)
