from fpdf import FPDF

comandos = [
    {
        "categoria": "Entrada e Saída (stdio.h)",
        "funcoes": [
            ("printf", "Imprime dados formatados.\nEx: printf(\"Idade: %d\", idade);"),
            ("scanf", "Lê dados da entrada padrão.\nEx: scanf(\"%d\", &idade);"),
            ("fgets", "Lê uma linha de texto com segurança.\nEx: fgets(nome, 100, stdin);"),
            ("puts", "Imprime uma string com quebra de linha.\nEx: puts(\"Olá\");"),
            ("putchar", "Imprime um caractere.\nEx: putchar('A');"),
        ]
    },
    {
        "categoria": "Conversão e Aleatoriedade (stdlib.h)",
        "funcoes": [
            ("atoi", "Converte string para inteiro.\nEx: int num = atoi(\"123\");"),
            ("rand", "Gera número aleatório.\nEx: int r = rand();"),
            ("srand", "Define semente para rand().\nEx: srand(time(NULL));"),
            ("exit", "Termina o programa.\nEx: exit(1);"),
            ("malloc", "Aloca memória dinamicamente.\nEx: int *p = malloc(sizeof(int));"),
        ]
    },
    {
        "categoria": "Tempo (time.h)",
        "funcoes": [
            ("time", "Retorna o tempo atual.\nEx: time_t t = time(NULL);"),
            ("difftime", "Calcula diferença entre dois tempos.\nEx: difftime(t2, t1);"),
            ("struct tm", "Estrutura para datas.\nEx: struct tm *info = localtime(&t);"),
        ]
    },
    {
        "categoria": "Strings (string.h)",
        "funcoes": [
            ("strlen", "Retorna o tamanho da string.\nEx: strlen(nome);"),
            ("strcpy", "Copia uma string.\nEx: strcpy(dest, src);"),
            ("strcat", "Concatena strings.\nEx: strcat(dest, src);"),
            ("strcmp", "Compara duas strings.\nEx: strcmp(s1, s2);"),
            ("strcspn", "Índice do 1º caractere de um conjunto.\nEx: strcspn(s, \"\\n\");"),
        ]
    },
    {
        "categoria": "Caracteres (ctype.h)",
        "funcoes": [
            ("isdigit", "Verifica se é dígito.\nEx: isdigit(c);"),
            ("isalpha", "Verifica se é letra.\nEx: isalpha(c);"),
            ("isspace", "Verifica se é espaço/tab.\nEx: isspace(c);"),
            ("toupper", "Converte para maiúscula.\nEx: toupper(c);"),
            ("tolower", "Converte para minúscula.\nEx: tolower(c);"),
        ]
    },
]

pdf = FPDF()
pdf.set_auto_page_break(auto=True, margin=15)
pdf.add_page()
pdf.set_font("Arial", "B", 16)
pdf.cell(0, 10, "Cheatsheet - Funções principais da linguagem C", ln=True, align="C")
pdf.ln(5)
pdf.set_font("Arial", "", 12)

for categoria in comandos:
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, categoria["categoria"], ln=True)
    pdf.set_font("Arial", "", 12)
    for nome, descricao in categoria["funcoes"]:
        pdf.cell(0, 8, f"{nome}:", ln=True)
        for linha in descricao.split("\n"):
            pdf.cell(0, 7, f"  {linha}", ln=True)
    pdf.ln(4)

pdf.output("utils/cheatsheet_c_funcoes.pdf")
