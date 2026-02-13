# Planejamento Wizped Office

Este pequeno projeto é um sistema de gerenciamento simplificado de estudantes de idiomas de uma escola Wizard by Pearson unidade de Naviraí.

## Casos de uso e funções

O sistema será colaborativo dado que será usado em mais de um PC ao mesmo tempo usando o GoogleDrive ou OneDrive como forma de sincronização. No entanto ainda há de se decidir a forma de armazenar os dados (banco de dados). A princípio está se usando as planilhas com suas tabelas na seguinte configurações:

> (planilha - tabela)

- BD_Alunos - Tbl_Alunos
- BD_Estagios - Tbl_Estagios
- BD_Modalidades - Tbl_Modalidade
- BD_TipoModalidade
- BD_Funcionarios - Tbl_Funcionarios
- BD_Status - Tbl_Status
- BD_Contrato
- BD_Motivos - Tbl_Motivos
- BD_TipoAula - Tbl_TipoAula
- BD_Horarios - Tbl_Horarios
- BD_Logs - Tbl_Logs

## Tabelas

### Tbl_Alunos

Esta tabela será onde a lista de alunos com seu nome completo será armazenada. Contém as colunas principais:

- ID_Aluno
- Nome

Essa tabela também contém as colunas de Estagios (livros), Modalidade de curso, Status, Contrato, TipoAula, Horarios e Logs.

Todas essas colunas são dados que podem ser editados na região/sessão do aluno ou escolhidos quando se cadastra um aluno novo, mas eles já são previamente cadastrados antes de criar um aluno novo ou editá-lo.

### Tbl_Estagios

Esta tabela armazena os livros disponíveis para um aluno usar. Ela tem as colunas:

- ID_Estagio
- Nome

Geralmente esta tabela não tem muita movimentação de dados já que a quantidade de livros costuma ser fixa, mas eventualmente pode haver modificações dos mesmos como troca do nome ou edições/versões diferentes, ou até adição/remoção de poucos livros.

### Tbl_Modalidades

Aqui se armazena algumas modalidades do curso dependendo do perfil de cada aluno. Costumamos usar as seguintes:

- Connections: Aulas padrão tradicional (costuma ser presencial)
- Interactive: Aulas com tablet (costuma ser presencial) ou sistema Catchup (o Catchup é online/wizardon)
- Online: Aulas online sem contrato WizardON
- WizardON: Aulas online com contrato WizardON

As colunas de Tbl_Modalidade são:

- ID_Modalidade
- Descricao

### Tbl_TipoModalides

Aqui são combinações/variações de cada modalide de cursos. Podemos combinar as modalidades. Primeiro temo o modo VIP, e então as combinações das outras:

- Connections (presencial)
- Connections VIP (presencial)
- Interactive (presencial)
- Interactive VIP (turma)
- Online VIP
- Connections Online
- Interactive Online
- Connections VIP Online
- Interactive VIP Online
- WizardON
- WizardON Connections
- WizardON Interactive
- WizardON VIP Connections
- WizardON VIP Interactive

> **Observação:** Em TipoModalide, quando um Aluno tem Online ou WizardON sozinhos como na lista acima, quer dizer que o aluno não está fazendo o curso padrão, ou seja, não usando o livro do curso e nem metodologia em si, mas sim um curso avulso ou paralelo para viajens, conversação ou algo do gênero.

Ainda há de revisar as combinações mais eficientes, flexíveis e/ou reaproveitáveis dado que haverão filtros por (tipo de aula):

- Se é presencial ou online (tanto Online padrão quanto WizardON)
- Se é Interactive ou Connections
- As combinações com VIP
- E as combinações variadas para cada caso

> A primeira hierarquia de filtro como termos mais umbrella, se usa Interactive, Connections e Online (seja online padrão ou WizardON). Depois disso vem a segunda parte e assim por dantes. Pode-se dizer essas palavras chaves são "tags" pois um aluno pode ser multicategorizadoTbl_Funcionarios

Colunas:

- ID_Funcionario
- Nome

Aqui vai os funcionários envolvidos, sejam da secretaria, professores, comercial e etc.

### Tbl_Status

Colunas

- ID_Status
- Descricao
- Sigla

Esta tabela possui o estado atual de um aluno, sendo:

- Ativo (A): quando o aluno está ativamente fazendo o estagio
- Desistente (D): quando um aluno para o estágio/livro sem concluir (é evasão)
- Encerrado (E): quando o livro/estagio for terminado
- Formado (F): quando o aluno finaliza todos os estagios
- Trancado (T): quando o aluno pausa o curso e retorna após um tempo determinado. Mas durante esse tempo esse trancamento pode se tornar em uma desistencia.

### Tbl_Motivos

Colunas:

- ID_Motivo
- Descricao

Esta tabela contém razões pré-cadastradas pelas quais um aluno esteja em um destes status:

- Desistente
- Trancado

Os motivos mais frequentes costumam ser:

- Atraso de pagamento
- Concorrência
- Desemprego
- Faltas frequentes
- Mudança de cidade
- Não se adaptou ao curso
- Questões pessoais
- Trabalho
- Viajem

### Tbl_Contrato

- ID_Contrato
- Tipo

Esta tabela armazena duas formas de contrato:

- Matrícula: quando um aluno é novo e inicia seu primeiro estágio
- Rematrícula: quando um aluno matricula novamente/continua/dá continuidade para outro/seguinte estágio, seja qualquer idioma.

### Tbl_TipoAula

Colunas:

- ID_TipoAula
- Descricao

Esta tabela possui tipos comuns de aula quando o aluno vem à aula. Quem registra é o professor. Os valores costumam ser:

- Normal
- Anteposição
- Reposição
- Reforço
- Diferenciada

### BD_Horarios

Colunas:

- ID_Horario
- Data
- Hora

Esta tabela é especial pois ela registra horários de alunos registrados. Vários alunos podem ter o mesmo horários dado que há turmas. 

Por exemplo (relato de uma gravação):

{
> Pedro costuma fazer suas aulas das terças e quintas, das 9h às 10h. Costuma ser assim. Mas ele já trocou o horário dele várias vezes. Ele já fez das 10h às 11h, ele já fez quarta e sexta, das 1h às 2h, ele já fez segunda e quarta, das 2h às 3h, ele já fez em dias diferentes, mas horários diferentes também, tipo assim, ele já fez terça, da 1h às 2h da tarde e quarta, das 3h às 4h da tarde. Ele já chegou a fazer três aulas também, por exemplo, segunda, 1h da tarde e terça-feira, da 1h às 2h e das 2h às 3h, ou seja, dois aulas seguidas, né? O aluno Vitor, ele também acontece isso, né? Ele é mais de boa, ele costuma fazer mais à noite. Ele geralmente faz terças e quintas, das 5h às 6h da tarde, fixo, ele é mais fixo, ele não muda tanto quanto o Pedro. Mas quando ele falta, ele costuma fazer reposição no sábado, ele faz das 9h às 11h da manhã, as dois aulas que ele perde. A Maria também, ela também tem um horário fixo, né? Por ela mesma, porque ela trabalha de manhã, né? Não todos os dias, mas ela faz as dois aulas dela, em vez de fazer dois dias diferentes, ela faz na sexta-feira, das 2h às 4h da tarde. Então os alunos eles podem ter horários bem flexíveis, entendeu? Não é rígido. Então existe uma múltipla combinação de horários. O que que é um horário? Horário é a junção da data e da hora, entendeu? E também muitas vezes, por exemplo, tem mais um aluno novo, né? O nome dele é João e o João ele também tem o mesmo horário do Pedro, ok? Porque eles estão na mesma turma. Esses são os exemplos claros de alguns alunos, né? Existem muito mais alunos com horários diferenciados. Então manter um traqueamento disso acaba sendo complicado de cabeça e na mão usando um papel. Então esse sistema precisa lidar com isso. Então por isso que essa tabela tem essas três colunas. Uma outra questão é o seguinte, em relação aos horários dos alunos, há de se observar e verificar como vai ser trabalhado de acordo com a normalização de dados, né, e boas práticas e convenções e diretivas de banco de dados. Considerando o ID, né, a chave primária de cada tabela, há de verificar como é que a tabela de horários vai se comunicar com a tabela de alunos, ok? Ela não tem relação com nenhuma outra tabela. Ela costuma ter relação direta com a tabela de alunos, ok? As outras tabelas, assim, são informações adicionais para acrescentar no aluno, como você pode ver nas descrições anteriores. Então, ficamos na dúvida se usar o ID próprio da tabela de horários ou seja o mesmo ID do aluno, como chave estrangeira. Ainda não sei como vou resolver isso. O que seria mais eficiente? Vai depender da estratégia de banco de dados, como eu vou falar sobre isso mais adiante.

}

### Tbl_Logs

Colunas:

- ID_Log
- Tabela
- Descricao
- Detalhes
- Data_hora (timestamp)

Esta tabela vai registrar cada movimento, cadastramento, alterações, edições, exclusões de algumas tabelas que envolvem diretamente o aluno.

Exemplo (relato de um áudio):

{
> esta tabela de logs é o seguinte, ela precisar registrar frequências todos os dias de ações de qualquer tipo. Eu disse anteriormente que pode ser até cadastramento, só que na verdade é mais ideal registrar ações como falta de aluno, presença, qual que aluno veio ou não veio, como eu falei antes, né? entrada de aluno, matrícula, rematrícula, troca de livro, troca de horário, entendeu? Cadastramento de aluno, exclusão de aluno, mudanças de status do aluno, né? Mudança de modalidade, ou seja, mudanças de alunos, eu posso fazer isso, entendeu? Talvez essa tabela precisa de mais colunas e talvez uma esquema melhor. Talvez eu precisasse fazer outra tabela, pudesse relacionar todas, né? Mas talvez, depende. Por exemplo, se eu vou acrescentar um livro, isso aí não tem necessidade de ter no log, né? Ah, eu vou adicionar um livro novo, vou deletar, vou alterar. Isso não tem necessidade. Os logs têm que, na verdade, estar relacionado a ações direcionadas aos alunos em si, não à empresa, né? Se, por exemplo, eu colocar ali, ah, vou adicionar uma modalidade nova, eu vou editar, excluir, né? Uma modalidade nova, um tipo de modalidade, um status, professores, funcionários, no caso. Tudo que é voltado à empresa não precisa de logs. É mais sim a coisas relacionadas ao aluno, entendeu? Ele que precisa. E o que mais vai mudar aqui é a questão de contrato, quando o aluno é matricular ou rematricular, né? Porque assim, isso é importante porque tem que gerar um timestamp, uma data e hora, que vai ser usada como filtro, entendeu? Pra gerar relatórios precisos. Entendeu? Porque na empresa nós manualmente temos tabelas impressas, né, que imprimimos aí na na Folha 4 de, por exemplo, alunos matriculados esse ano, por exemplo, né? A gente coloca, escreve sempre quando o aluno novo começa, a data que ele começou, né, e qual livro foi entregado por esse aluno, né, que é a entrega de materiais, no caso. Geralmente é alunos novos, que é matrícula, no caso, e entrega de materiais meio que andam juntos, entendeu? Às vezes acontece o aluno matricular primeiro, só que receber o livro depois, então isso aí geralmente a gente faz em em folhas diferentes, entendeu? Aí a gente sempre contabiliza manualmente e isso acaba gerando um esforço logístico muito grande, entendeu? Tendo que revisar várias vezes, porque a mente humana, ela é falha. Então, talvez, criar uma outra tabela que pudesse relacionar uma tabela unificada, entendeu? Talvez. com colunas diferentes. Eu não sei como é que eu vou fazer ainda esse esquema de logs, entendeu? Mas outra coisa é alunos assistentes, alunos que rematricularam, é outras folhas que nós imprimimos, no caso, são listas simples, né, tabelas simples que eu imprimo ali no Excel e imprimo numa Folha 4. Eu quero substituir tudo isso em um, tipo uma tabela de logs, entendeu? Que seria muito mais eficiente. Então, essa é a situação que eu quero eficientizar, né, e ser leve e rápido, usando estratégias aí de convenções boas de banco de dados, só que no Excel. ou qualquer outra forma que fosse eficiente sem quebrar o projeto e o sistema. por exemplo, uma situação, é se aonde ele é dependendo do status dele, vamos supor, ele é trancado, né? Aí qual que é o motivo do trancamento dele? Foi viagem ou foi questão financeira dele, porque ele não estava pagando, ou seja lá o que for, né? E por e também desistência, né? Se ele desistiu por quê? Mudou de cidade, né? Por que que houve uma evasão no total? Eu pensei tipo assim, colocar ali o status e depois o tipo do status, né, que é a combinação de outras coisas, né? ou não, né, na verdade. Enfim, não sei como é que eu vou fazer, porque nós temos a a coluna ali, né? de descrição e detalhes. Eu não sei o que que eu vou colocar em cada uma, né? Ah. Se dá pra concatenar tudo, talvez, a gente coloca ali concatenado alguns valores de colunas, né? Talvez fosse melhor em uma coluna apenas um registro de evento em uma string pura ou não, aí lá no front-end ou aqui na na interpretação aqui do Excel, eu posso colocar o texto puro, sei lá, eu não sei o que que seria mais eficiente. Tá, vai depender da estratégia de banco de dados, como eu como eu foi mencionado anteriormente, né? Como é que a gente como é que registramos ações, né? Por exemplo, a minha colega na recepção vai registrar uma falta de aluno e vai aparecer pra mim automaticamente lá na no andar de cima, na planilha também. Talvez em planilha em passa de trabalho diferente do Excel, só que usando a mesma fonte do Google Drive, por exemplo, ou OneDrive, tanto faz isso aí. né? Talvez arquivos em comum ali, um arquivo de dados que podem compartilhar no caso, né? É, eu não sei como é que vai ser isso, porque a estratégia inicial que eu pensei em questão de atualização de banco de dados compartilhado, seria The last the last one wins. É tipo assim, a a alteração mais recente é a que vai valer. Entendeu? E outra coisa importante é colocar quem que foi o autor dessa manipulação, seja cada- qualquer operação de CRUD, entendeu? Ah, geralmente é atualização de aluno, de alguma coisa, frequência vai ser mais vezes. É vanterco, rematrícula, entrega de material, etc. Entendeu? Ah, eu pe- na verdade, assim, em todo esse projeto, lá do início que eu comecei a te falar sobre isso, né? Eu estava pensando em usar o Access, no caso, pra isso. Entendeu? Porque eu não que- eu quero criar um projeto portátil. Porque assim, como os dois vão usar uma sincroni- sincronização via Google Drive, eu tenho que ter cuidado com isso também. Como é que eu posso fazer isso, entendeu? Talvez uma replicação diferente, usando o VBA, porque que eu vou eu vou ter que usar o VBA por trás disso aqui, né? E fazer certinho, porque, por exemplo, o PowerPivot do do Excel, eh, eu tenho que ver o que que faz sentido usar nele. Porque assim, as tabelas eu posso transformá-las em modelos de dados, com certeza e fazer relações com o banco de dados normal. Eu abro aqui a interface e eu consigo puxar uma linha, é tipo assim, fazer um link, né, entre as colunas e correspondente, né? De com outra com outra tabela, etc. Eu já fiz isso em algumas questões. Você pode ver na imagem ali, eh como eu fiz isso. Aquilo lá é uma versão só de protótipo, não é definitiva, é apenas estou testando coisas. Mas há de avaliar se eu é, talvez seja melhor até usar o Access, mas aí eu posso usar todo o sistema no Access, porque assim, é a questão. Eu quero usar o Excel porque eu preciso imprimir certas coisas e templates, né, de de de certas coisas. Eu preciso fazer isso. Né? Então, o caminho é grande ainda.

}

### Detalhes importantes

Relato de um áudio:

{
> Um outro detalhe importante, que é o seguinte, na tabela de logs, né, por exemplo, um aluno está cursando o livro, supor que ele faz 70% do livro e ele para o curso, ele tranca, aí beleza, ele tem um trancamento por um motivo específico, sei lá, viagem, né, ou sei lá, questões pessoais, ou questão financeira. Enfim, isso não vem ao caso, mas ele pode, tipo assim, trancar, aí ele volta, aí isso precisa estar no log, entendeu? Né, Então a gente coloca ele de status de trancado para ativo. Ativo é destrancado, é a mesma coisa. Então vamos colocar ativo, né? Ah, isso é importante. Em questão dos estágios, que são os livros praticamente, né, um aluno ele pode estar em mais de um livro ao mesmo tempo. Claro, não no mesmo horário, né, com certeza. É impossível o aluno estar no mesmo livro no mesmo horário. Tipo assim, vamos supor que o aluno Pedro ele faz um livro de inglês num horário e o espanhol em outro, né? Ah, Tipo assim, não tem como ele fazer os dois livros ao mesmo tempo. Isso não é, isso é inválido. Então, por exemplo, o contrato que tem dois tipos, né, que é matrícula e rematrícula. Ah, vamos supor que o Pedro está no segundo livro de inglês, só que no primeiro livro de espanhol. Então ele é rematrícula, considerando o atual livro de inglês dele, né, e é matrícula no espanhol porque é aluno novo nesse nesse nesse primeiro livro de espanhol, entendeu? Então isso conta como matrículas. A matrícula se configura praticamente quando o aluno começa o primeiro livro. Entendeu? Da Wizard. Não necessariamente, mas um primeiro livro. Mas é um aluno que inicia um curso que ele não fez antes. Vamos supor que vem um aluno de fora, sei lá, a Catarina. Ela começa o livro de espanhol e é matrícula, porque ela tá começando a estudar com a gente, é aluna nova. Mas o Pedro não. Mas se o Pedro já fez inglês em uma vez e ele tá no segundo módulo, ou segundo estágio, ele começou a fazer espanhol, isso também é considerado matrícula porque ele começou o livro de espanhol. Ou seja, quando o aluno inicia um módulo, o primeiro, no caso o estágio, né, é configurado matrícula. Tá? E outra questão importante é que um aluno, ele pode estar vinculado a múltiplos professores ao mesmo tempo, na mesma sala de aula, entendeu? Eu não quero tabelas de sala de aula. Isso é irrelevante para a gente. Mas, por exemplo, geralmente eu fico eu e o professor Williams na sala de aula. Às vezes a professora Maria ajuda também, porque a sala acaba sendo grande, entendeu? É uma sala gigante. Parece um mini auditório. Então a gente lida com vários alunos. Então, podemos fazer isso também. Bom, em questão de multitarefas simultâneas, o aluno ele pode cursar curso de inglês ou espanhol ao mesmo tempo. Sim, ele deve, o seu formulário deve permitir eu gerenciar todas as matérias ativas de um aluno, né, na tela, ou prefere trabalhar matrícula por vez. Na verdade, assim, quando o aluno ele faz uma matrícula, ou seja, um contrato, né, do tipo matrícula, é porque ele está começando o curso, como eu te falei antes. E sim, o formulário deve permitir gerenciar todas as matrículas ativas, ou rematrículas também na mesma tela. Na verdade, assim, no sistema que nós temos já interno da Wizard, ele tem um código, né, que é o código de contrato barra e o número da matrícula. Tipo assim, o aluno não está em dois livros, o número do contrato do aluno, que é um número inteiro, é de geralmente 4 dígitos, né, barra e qual que é o livro. Isso junto, eu não... essa sequência é o número de contrato, como funciona aqui, entendeu? Eu não quero usar isso no meu sistema, mas é só pra você entender como é que funciona. É o mesmo contrato, só que as matrículas diferentes. Por quê? Porque esse, por exemplo, 2000 barra 1 é o aluno número 2000, não quer dizer que é 2.000 alunos, tá? É só o número que o sistema gera, número inteiro, né? E o número 1 é o primeiro livro do aluno, entendeu? Ou o primeiro que foi lançado, na verdade, né? Um sistema matriculado pra ele o aluno estudar. E o 2000 barra 2 é o outro livro do aluno. Pode ser o outro livro que ele está fazendo o curso ao mesmo tempo, ou pode ser o seguinte livro do curso dele. do que vem depois do estágio que ele finalizou. Isso aí pode acontecer, entendeu? Então, eu acredito que talvez seja até melhor separar, entendeu? É isso aí, para gerenciar a matrícula ou o tipo de contrato. O contrato, na verdade, ele é matrícula, né? Ele tem dois tipos. E matrícula, na verdade, tem o livro, no caso o estágio, o aluno quando começou ou não, quando vai finalizar, et cetera. Em questão do Excel e o Access, se eu estou indeciso ou não, em questão da arquitetura, ela muda bastante, você falou, né? Ah, eu acho que eu vou usar tipo três planilhas, três pastas de trabalho do Excel. Duas com macro e uma sem. A padrão XLSX eu vou usar como os dados. Lá que vai ficar todo o banco de dados, entendeu? Eu não vou usar o Access porque ele dá lock no banco de dados para poder usar em vários lugares, porque eu vou sincronizar via OneDrive, entendeu? Então, para mim, é mais fácil fazer usar um XLSX, no final das contas, para gerar todo o esquema, né? E os outros dois, eu... eu vou, um vou usar aqui no meu computador de cima, no notebook, e outro vai ser usado no computador da recepção. Aí vai ser dois workbooks, um de macro, um pedagógico e outro administrativo, mas eles vão se comunicar entre si, entendeu? Eu acredito que seja mais eficiente assim. Bom, eu suponho. Porque se eu abrir o mesmo arquivo de Excel... Do OneDrive em cada computador vai dar conflito, entendeu? Vai dar um pouco de conflito, eu acho. Eu pensei até usar o mesmo arquivo, um só, mas acho que vai dar problema. Ah, o escopo do formulário é em questão de verificar estruturar os campos necessários do formulário do aluno, né? Na verdade, eu quero que você foque mais em como criar os componentes dentro do Excel, né? Os formulários, os componentes de VBA e etc. E no final a gente faz a geração da automática da ficha de frequência, após tudo estar funcionando, entendeu? Porque senão fica muito complexo, mas seria mais assim, seria bom gerar lá uma coisa bem parecida. Porque eu já tenho uma planilha que eu vou usar no meu sistema, a parte administrativa, no recepção, vai, ela já vai estar aceita, na verdade, a tabela. Só que eu quero um VBA que gere ela dependendo da estratégia que eu vou colocar depois. Eu quero focar no esquema geral, entendeu? Pra depois começar a fazer as telas e depois fazer as automações. Pensei também em usar uma ribbon customizada para cada workbook se necessário.
Para testar tenho os arquivos totalmente limpos e vazios de dados e códigos: BD_Wiz_Dados.xlsx ("C:\Users\Vitor\Documents\wizped\BD_Wiz_Dados.xlsx"), APP_Wizped.xlsm ("C:\Users\Vitor\Documents\wizped\APP_Wizped.xlsm ") e APP_Wiz_admin.xlsm ("C:\Users\Vitor\Documents\wizped\APP_Wiz_admin.xlsm").

}

## Componentes Excel VBA

> Abaixo está o mapeameno de vários componentes do Excel VBA com valores padrão

### Label (Rótulo)

> Label: Controle gráfico utilizado para exibir texto estático que o usuário não pode editar diretamente. É empregado para rotular outros controles, fornecer instruções ou exibir informações de saída somente leitura na interface.

| Categoria | Propriedade | Valor |
| :--- | :--- | :--- |
| **Aparência** | (Name) | Label1 |
| | BackColor | &H8000000F& |
| | BackStyle | 1 - fmBackStyleOpaque |
| | BorderColor | &H80000006& |
| | BorderStyle | 0 - fmBorderStyleNone |
| | Caption | Label1 |
| | ControlTipText | |
| | ForeColor | &H80000012& |
| | SpecialEffect | 0 - fmSpecialEffectFlat |
| | Visible | True |
| **Comportamento** | AutoSize | False |
| | Enabled | True |
| | TextAlign | 1 - fmTextAlignLeft |
| | WordWrap | True |
| **Diversos** | Accelerator | |
| | HelpContextID | 0 |
| | MouseIcon | (Nenhum) |
| | MousePointer | 0 - fmMousePointerDefault |
| | TabIndex | 23 |
| | TabStop | False |
| | Tag | |
| **Fonte** | Font | Tahoma |
| **Imagem** | Picture | (Nenhum) |
| | PicturePosition | 7 - fmPicturePositionAboveCenter |
| **Posição** | Height | 18 |
| | Left | 72 |
| | Top | 282 |
| | Width | 72 |

### Text Box (Caixa de Texto)

> TextBox: Campo de entrada de dados que permite ao usuário digitar, visualizar ou editar cadeias de texto. Suporta funcionalidades como máscaras de senha, múltiplas linhas de texto e barras de rolagem internas.

| Categoria | Propriedade | Valor |
| :--- | :--- | :--- |
| **Aparência** | (Name) | TextBox1 |
| | BackColor | &H80000005& |
| | BackStyle | 1 - fmBackStyleOpaque |
| | BorderColor | &H80000006& |
| | BorderStyle | 0 - fmBorderStyleNone |
| | ControlTipText | |
| | ForeColor | &H80000008& |
| | PasswordChar | |
| | SpecialEffect | 2 - fmSpecialEffectSunken |
| | Value | |
| | Visible | True |
| **Comportamento** | AutoSize | False |
| | AutoTab | False |
| | AutoWordSelect | True |
| | Enabled | True |
| | EnterKeyBehavior | False |
| | HideSelection | True |
| | IntegralHeight | True |
| | Locked | False |
| | MaxLength | 0 |
| | MultiLine | False |
| | SelectionMargin | True |
| | TabKeyBehavior | False |
| | TextAlign | 1 - fmTextAlignLeft |
| | WordWrap | True |
| **Dados** | ControlSource | |
| | Text | |
| **Diversos** | DragBehavior | 0 - fmDragBehaviorDisabled |
| | EnterFieldBehavior | 0 - fmEnterFieldBehaviorSelectAll |
| | HelpContextID | 0 |
| | IMEMode | 0 - fmIMEModeNoControl |
| | MouseIcon | (Nenhum) |
| | MousePointer | 0 - fmMousePointerDefault |
| | TabIndex | 23 |
| | TabStop | True |
| | Tag | |
| **Fonte** | Font | Tahoma |
| **Posição** | Height | 18 |
| | Left | 108 |
| | Top | 276 |
| | Width | 72 |
| **Rolagem** | ScrollBars | 0 - fmScrollBarsNone |

### Combo Box (Caixa de Combinação)

> ComboBox: Controle que combina uma caixa de texto com uma lista suspensa. Permite ao usuário selecionar um item de uma lista pré-definida ou, dependendo da configuração, inserir um valor personalizado diretamente no campo.

| Categoria | Propriedade | Valor |
| :--- | :--- | :--- |
| **Aparência** | (Name) | ComboBox1 |
| | BackColor | &H80000005& |
| | BackStyle | 1 - fmBackStyleOpaque |
| | BorderColor | &H80000006& |
| | BorderStyle | 0 - fmBorderStyleNone |
| | ControlTipText | |
| | DropButtonStyle | 1 - fmDropButtonStyleArrow |
| | ForeColor | &H80000008& |
| | ShowDropButtonWhen | 2 - fmShowDropButtonWhenAlways |
| | SpecialEffect | 2 - fmSpecialEffectSunken |
| | Style | 0 - fmStyleDropDownCombo |
| | Value | |
| | Visible | True |
| **Comportamento** | AutoSize | False |
| | AutoTab | False |
| | AutoWordSelect | True |
| | Enabled | True |
| | HideSelection | True |
| | Locked | False |
| | MatchEntry | 1 - fmMatchEntryComplete |
| | MatchRequired | False |
| | MaxLength | 0 |
| | SelectionMargin | True |
| | TextAlign | 1 - fmTextAlignLeft |
| **Dados** | BoundColumn | 1 |
| | ColumnCount | 1 |
| | ColumnHeads | False |
| | ColumnWidths | |
| | ControlSource | |
| | ListRows | 8 |
| | ListStyle | 0 - fmListStylePlain |
| | ListWidth | 0 pt |
| | RowSource | |
| | Text | |
| | TextColumn | -1 |
| | TopIndex | -1 |
| **Diversos** | DragBehavior | 0 - fmDragBehaviorDisabled |
| | EnterFieldBehavior | 0 - fmEnterFieldBehaviorSelectAll |
| | HelpContextID | 0 |
| | IMEMode | 0 - fmIMEModeNoControl |
| | MouseIcon | (Nenhum) |
| | MousePointer | 0 - fmMousePointerDefault |
| | TabIndex | 23 |
| | TabStop | True |
| | Tag | |
| **Fonte** | Font | Tahoma |
| **Posição** | Height | 18 |
| | Left | 180 |
| | Top | 258 |
| | Width | 72 |

### Lista Box (Caixa de listagem)

> ListBox: Exibe uma lista de itens de onde o usuário pode selecionar um ou mais elementos. Diferente do ComboBox, a lista permanece visível e a seleção pode ser configurada para ser única ou múltipla.

| Categoria | Propriedade | Valor |
| :--- | :--- | :--- |
| **Aparência** | (Name) | ListBox1 |
| | BackColor | &H80000005& |
| | BorderColor | &H80000006& |
| | BorderStyle | 0 - fmBorderStyleNone |
| | ControlTipText | |
| | ForeColor | &H80000008& |
| | SpecialEffect | 2 - fmSpecialEffectSunken |
| | Value | |
| | Visible | True |
| **Comportamento** | Enabled | True |
| | IntegralHeight | True |
| | Locked | False |
| | MatchEntry | 0 - fmMatchEntryFirstLetter |
| | MultiSelect | 0 - fmMultiSelectSingle |
| | TextAlign | 1 - fmTextAlignLeft |
| **Dados** | BoundColumn | 1 |
| | ColumnCount | 1 |
| | ColumnHeads | False |
| | ColumnWidths | |
| | ControlSource | |
| | ListStyle | 0 - fmListStylePlain |
| | RowSource | |
| | Text | |
| | TextColumn | -1 |
| | TopIndex | -1 |
| **Diversos** | HelpContextID | 0 |
| | IMEMode | 0 - fmIMEModeNoControl |
| | MouseIcon | (Nenhum) |
| | MousePointer | 0 - fmMousePointerDefault |
| | TabIndex | 23 |
| | TabStop | True |
| | Tag | |
| **Fonte** | Font | Tahoma |
| **Posição** | Height | 72 |
| | Left | 228 |
| | Top | 288 |
| | Width | 72 |

### Check Box (Caixa de Seleção)

> CheckBox: Fornece uma opção de alternância binária (Verdadeiro/Falso). É utilizado para permitir que o usuário faça escolhas independentes, onde múltiplas opções em um grupo podem ser selecionadas simultaneamente.

| Categoria | Propriedade | Valor |
| :--- | :--- | :--- |
| **Aparência** | (Name) | CheckBox1 |
| | Alignment | 1 - fmAlignmentRight |
| | BackColor | &H8000000F& |
| | BackStyle | 1 - fmBackStyleOpaque |
| | Caption | CheckBox1 |
| | ControlTipText | |
| | ForeColor | &H80000012& |
| | SpecialEffect | 2 - fmButtonEffectSunken |
| | Value | False |
| | Visible | True |
| **Comportamento** | AutoSize | False |
| | Enabled | True |
| | Locked | False |
| | TextAlign | 1 - fmTextAlignLeft |
| | TripleState | False |
| | WordWrap | True |
| **Dados** | ControlSource | |
| **Diversos** | Accelerator | |
| | GroupName | |
| | HelpContextID | 0 |
| | MouseIcon | (Nenhum) |
| | MousePointer | 0 - fmMousePointerDefault |
| | TabIndex | 23 |
| | TabStop | True |
| | Tag | |
| **Fonte** | Font | Tahoma |
| **Imagem** | Picture | (Nenhum) |
| | PicturePosition | 7 - fmPicturePositionAboveCenter |
| **Posição** | Height | 18 |
| | Left | 126 |
| | Top | 270 |
| | Width | 108 |

### Option Button (Botão de Opção)

> OptionButton: Também conhecido como botão de rádio, permite a seleção de uma única opção dentro de um grupo. Quando um botão é ativado, os outros botões do mesmo grupo ou container são automaticamente desativados.

| Categoria | Propriedade | Valor |
| :--- | :--- | :--- |
| **Aparência** | (Name) | OptionButton1 |
| | Alignment | 1 - fmAlignmentRight |
| | BackColor | &H8000000F& |
| | BackStyle | 1 - fmBackStyleOpaque |
| | Caption | OptionButton1 |
| | ControlTipText | |
| | ForeColor | &H80000012& |
| | SpecialEffect | 2 - fmButtonEffectSunken |
| | Value | False |
| | Visible | True |
| **Comportamento** | AutoSize | False |
| | Enabled | True |
| | Locked | False |
| | TextAlign | 1 - fmTextAlignLeft |
| | TripleState | False |
| | WordWrap | True |
| **Dados** | ControlSource | |
| **Diversos** | Accelerator | |
| | GroupName | |
| | HelpContextID | 0 |
| | MouseIcon | (Nenhum) |
| | MousePointer | 0 - fmMousePointerDefault |
| | TabIndex | 23 |
| | TabStop | True |
| | Tag | |
| **Fonte** | Font | Tahoma |
| **Imagem** | Picture | (Nenhum) |
| | PicturePosition | 7 - fmPicturePositionAboveCenter |
| **Posição** | Height | 18 |
| | Left | 138 |
| | Top | 264 |
| | Width | 108 |

### Toggle Button (Botão de Ativação)

> ToggleButton: Um botão que alterna entre dois estados visuais (pressionado ou não pressionado). Funciona como um interruptor para ativar ou desativar uma função específica, mantendo o estado até ser clicado novamente.

| Categoria | Propriedade | Valor |
| :--- | :--- | :--- |
| **Aparência** | (Name) | ToggleButton1 |
| | BackColor | &H8000000F& |
| | BackStyle | 1 - fmBackStyleOpaque |
| | Caption | ToggleButton1 |
| | ControlTipText | |
| | ForeColor | &H80000012& |
| | Value | False |
| | Visible | True |
| **Comportamento** | AutoSize | False |
| | Enabled | True |
| | Locked | False |
| | TextAlign | 2 - fmTextAlignCenter |
| | TripleState | False |
| | WordWrap | True |
| **Dados** | ControlSource | |
| **Diversos** | Accelerator | |
| | HelpContextID | 0 |
| | MouseIcon | (Nenhum) |
| | MousePointer | 0 - fmMousePointerDefault |
| | TabIndex | 23 |
| | TabStop | True |
| | Tag | |
| **Fonte** | Font | Tahoma |
| **Imagem** | Picture | (Nenhum) |
| | PicturePosition | 7 - fmPicturePositionAboveCenter |
| **Posição** | Height | 40 |
| | Left | 210 |
| | Top | 264 |
| | Width | 36 |

### Frame (Quadro)

>  Frame: Atua como um container visual e funcional para agrupar controles relacionados. É essencial para organizar OptionButtons em grupos distintos e para gerenciar a visibilidade ou o estado de ativação de múltiplos elementos de uma vez.

| Categoria | Propriedade | Valor |
| :--- | :--- | :--- |
| **Aparência** | (Name) | Frame1 |
| | BackColor | &H8000000F& |
| | BorderColor | &H80000012& |
| | BorderStyle | 0 - fmBorderStyleNone |
| | Caption | Frame1 |
| | ControlTipText | |
| | ForeColor | &H80000012& |
| | SpecialEffect | 3 - fmSpecialEffectEtched |
| | Visible | True |
| **Comportamento** | Cycle | 0 - fmCycleAllForms |
| | Enabled | True |
| **Diversos** | HelpContextID | 0 |
| | MouseIcon | (Nenhum) |
| | MousePointer | 0 - fmMousePointerDefault |
| | TabIndex | 23 |
| | TabStop | True |
| | Tag | |
| | Zoom | 100 |
| **Fonte** | Font | Tahoma |
| **Imagem** | Picture | (Nenhum) |
| | PictureAlignment | 2 - fmPictureAlignmentCenter |
| | PictureSizeMode | 0 - fmPictureSizeModeClip |
| | PictureTiling | False |
| **Posição** | Height | 144 |
| | Left | 174 |
| | Top | 264 |
| | Width | 216 |
| **Rolagem** | KeepScrollBarsVisible | 3 - fmScrollBarsBoth |
| | ScrollBars | 0 - fmScrollBarsNone |
| | ScrollHeight | 0 |
| | ScrollLeft | 0 |
| | ScrollTop | 0 |
| | ScrollWidth | 0 |

### Tab Strip

> TabStrip: Controle de navegação composto por uma série de abas. Diferente do MultiPage, ele não possui páginas internas para controles; é usado para alterar o contexto ou filtrar dados de outros controles externos com base na aba selecionada.

| Categoria | Propriedade | Valor |
| :--- | :--- | :--- |
| **Aparência** | (Name) | TabStrip1 |
| | BackColor | &H8000000F& |
| | ControlTipText | |
| | ForeColor | &H80000012& |
| | Style | 0 - fmTabStyleTabs |
| | TabOrientation | 0 - fmTabOrientationTop |
| | Visible | True |
| **Comportamento** | Enabled | True |
| **Diversos** | HelpContextID | 0 |
| | MouseIcon | (Nenhum) |
| | MousePointer | 0 - fmMousePointerDefault |
| | TabIndex | 23 |
| | TabStop | True |
| | Tag | |
| **Fonte** | Font | Tahoma |
| **Guias** | MultiRow | False |
| | TabFixedHeight | 0 |
| | TabFixedWidth | 0 |
| | Value | 0 |
| **Posição** | Height | 108 |
| | Left | 240 |
| | Top | 270 |
| | Width | 144 |

### Multi Page (Multi-Página)

> MultiPage: Container de múltiplas páginas (objetos Page) acessíveis por guias. Cada página pode conter seu próprio conjunto de controles independentes, sendo ideal para organizar formulários complexos em seções distintas.

| Categoria | Propriedade | Valor |
| :--- | :--- | :--- |
| **Aparência** | (Name) | MultiPage1 |
| | BackColor | &H8000000F& |
| | ControlTipText | |
| | ForeColor | &H80000012& |
| | Style | 0 - fmTabStyleTabs |
| | TabOrientation | 0 - fmTabOrientationTop |
| | Visible | True |
| **Comportamento** | Enabled | True |
| **Diversos** | HelpContextID | 0 |
| | TabIndex | 23 |
| | TabStop | True |
| | Tag | |
| | Value | 0 |
| **Fonte** | Font | Tahoma |
| **Guias** | MultiRow | False |
| | TabFixedHeight | 0 |
| | TabFixedWidth | 0 |
| **Posição** | Height | 108 |
| | Left | 192 |
| | Top | 264 |
| | Width | 144 |

### Scroll Bar (Barra de rolagem)

> ScrollBar: Barra de rolagem autônoma usada para navegar por um intervalo de valores ou deslocar o conteúdo de uma área de visualização. Permite o ajuste fino através de um cursor que se move entre valores mínimos e máximos definidos.

| Categoria | Propriedade | Valor |
| :--- | :--- | :--- |
| **Aparência** | (Name) | ScrollBar1 |
| | BackColor | &H8000000F& |
| | ControlTipText | |
| | ForeColor | &H80000012& |
| | Orientation | -1 - fmOrientationAuto |
| | ProportionalThumb | True |
| | Value | 0 |
| | Visible | True |
| **Comportamento** | Enabled | True |
| **Dados** | ControlSource | |
| **Diversos** | HelpContextID | 0 |
| | MouseIcon | (Nenhum) |
| | MousePointer | 0 - fmMousePointerDefault |
| | TabIndex | 23 |
| | TabStop | True |
| | Tag | |
| **Posição** | Height | 63,8 |
| | Left | 186 |
| | Top | 240 |
| | Width | 12,75 |
| **Rolagem** | Delay | 50 |
| | LargeChange | 1 |
| | Max | 32767 |
| | Min | 0 |
| | SmallChange | 1 |

### Spin Button (Botão de rotação)

> SpinButton: Controle incremental composto por setas de aumento e diminuição. Geralmente é vinculado a um TextBox para permitir que o usuário altere valores numéricos em passos específicos (SmallChange) de forma rápida.

| Categoria | Propriedade | Valor |
| :--- | :--- | :--- |
| **Aparência** | (Name) | SpinButton1 |
| | BackColor | &H8000000F& |
| | ControlTipText | |
| | ForeColor | &H80000012& |
| | Orientation | -1 - fmOrientationAuto |
| | Value | 0 |
| | Visible | True |
| **Comportamento** | Enabled | True |
| **Dados** | ControlSource | |
| **Diversos** | HelpContextID | 0 |
| | MouseIcon | (Nenhum) |
| | MousePointer | 0 - fmMousePointerDefault |
| | TabIndex | 23 |
| | TabStop | True |
| | Tag | |
| **Posição** | Height | 25,5 |
| | Left | 150 |
| | Top | 246 |
| | Width | 12,75 |
| **Rolagem** | Delay | 50 |
| | Max | 100 |
| | Min | 0 |
| | SmallChange | 1 |

### Image (Imagem)

> Image: Controle dedicado à exibição de arquivos gráficos (bitmaps, ícones ou meta-arquivos). Suporta o redimensionamento da imagem para caber no controle (SizeMode) e o alinhamento do conteúdo visual na interface.

| Categoria | Propriedade | Valor |
| :--- | :--- | :--- |
| **Aparência** | (Name) | Image1 |
| | BackColor | &H8000000F& |
| | BackStyle | 1 - fmBackStyleOpaque |
| | BorderColor | &H80000006& |
| | BorderStyle | 1 - fmBorderStyleSingle |
| | ControlTipText | |
| | SpecialEffect | 0 - fmSpecialEffectFlat |
| | Visible | True |
| **Comportamento** | AutoSize | False |
| | Enabled | True |
| **Diversos** | MouseIcon | (Nenhum) |
| | MousePointer | 0 - fmMousePointerDefault |
| | Tag | |
| **Imagem** | Picture | (Nenhum) |
| | PictureAlignment | 2 - fmPictureAlignmentCenter |
| | PictureSizeMode | 0 - fmPictureSizeModeClip |
| | PictureTiling | False |
| **Posição** | Height | 72 |
| | Left | 192 |
| | Top | 264 |
| | Width | 72 |

### RefEdit

> RefEdit: Controle especializado para o ambiente Excel que permite ao usuário selecionar uma referência de célula ou intervalo (Range) diretamente na planilha. Durante a seleção, o controle minimiza temporariamente o formulário para facilitar a visualização das células e retorna o endereço do intervalo selecionado como uma string formatada.

| Categoria | Propriedade | Valor |
| :--- | :--- | :--- |
| **Aparência** | (Name) | RefEdit1 |
| | BackColor | &H80000005& |
| | BackStyle | 1 - fmBackStyleOpaque |
| | BorderColor | &H80000006& |
| | BorderStyle | 0 - fmBorderStyleNone |
| | ControlTipText | |
| | ForeColor | &H80000008& |
| | PasswordChar | |
| | SpecialEffect | 2 - fmSpecialEffectSunken |
| | Value | |
| | Visible | True |
| **Comportamento** | AutoSize | False |
| | AutoTab | False |
| | AutoWordSelect | False |
| | Enabled | True |
| | EnterKeyBehavior | False |
| | HideSelection | True |
| | IntegralHeight | True |
| | Locked | False |
| | MaxLength | 0 |
| | MultiLine | False |
| | SelectionMargin | True |
| | TabKeyBehavior | False |
| | TextAlign | 1 - fmTextTextAlignLeft |
| | WordWrap | True |
| **Dados** | Text | |
| **Diversos** | DragBehavior | 0 - fmDragBehaviorDisabled |
| | EnterFieldBehavior | 0 - fmEnterFieldBehaviorSelectAll |
| | HelpContextID | 0 |
| | IMEMode | 0 - fmIMEModeNoControl |
| | MouseIcon | (Nenhum) |
| | MousePointer | 0 - fmMousePointerDefault |
| | SelLength | 0 |
| | SelStart | 0 |
| | SelText | |
| | TabIndex | 23 |
| | TabStop | True |
| | Tag | |
| **Fonte** | Font | Tahoma |
| **Posição** | Height | 18 |
| | Left | 510 |
| | Top | 246 |
| | Width | 72 |
| **Rolagem** | ScrollBars | 0 - fmScrollBarsNone |

### Button Bar

> ButtonBar: Um controle de interface personalizado (frequentemente de bibliotecas ActiveX de terceiros) que agrupa múltiplos botões de ação em uma única estrutura horizontal ou vertical. É utilizado para consolidar comandos relacionados, como controles de reprodução de mídia ou ferramentas de edição, otimizando o espaço no formulário.

| Categoria | Propriedade | Valor |
| :--- | :--- | :--- |
| **Aparência** | (Name) | ButtonBar1 |
| | ControlTipText | |
| | Visible | True |
| **Diversos** | HelpContextID | 0 |
| | TabIndex | 23 |
| | TabStop | True |
| | Tag | |
| **Posição** | Height | 17,25 |
| | Left | 600 |
| | Top | 246 |
| | Width | 210 |

### Image combo

> ImageCombo: Uma variação do ComboBox que suporta a exibição de imagens para cada item da lista. Para funcionar plenamente, ele deve ser associado a um controle ImageList, que armazena os ícones. É ideal para interfaces que exigem identificação visual rápida, como seletores de status, tipos de arquivos ou categorias com ícones representativos.

| Categoria | Propriedade | Valor |
| :--- | :--- | :--- |
| **Aparência** | (Name) | ImageCombo1 |
| | BackColor | &H80000005& |
| | ControlTipText | |
| | ForeColor | &H80000008& |
| | Indentation | 0 |
| | Text | ImageCombo1 |
| | Visible | True |
| **Comportamento** | Enabled | True |
| | Locked | False |
| | OLEDDragMode | 0 - ccOLEDragManual |
| | OLEDDropMode | 0 - ccOLEDropNone |
| **Diversos** | (Personalizado) | |
| | (Sobre) | |
| | HelpContextID | 0 |
| | MouseIcon | (None) |
| | MousePointer | 0 - ccDefault |
| | TabIndex | 23 |
| | TabStop | True |
| | Tag | |
| **Fonte** | Font | Tahoma |
| **Posição** | Height | 15,75 |
| | Left | 540 |
| | Top | 102 |
| | Width | 114 |

### Image List

> ImageList: Controle não visual em tempo de execução que atua como um repositório central de imagens (ícones ou bitmaps). Ele armazena uma coleção de imagens que podem ser referenciadas por outros controles, como o ImageCombo, ListView ou TreeView, através de um índice ou chave. É fundamental para manter a consistência visual e otimizar o gerenciamento de recursos gráficos em interfaces complexas.

| Categoria | Propriedade | Valor |
| :--- | :--- | :--- |
| **Aparência** | (Name) | ImageList1 |
| **Comportamento** | MaskColor | &H00C0C0C0& |
| | UseMaskColor | True |
| **Diversos** | (Personalizado) | |
| | (Sobre) | |
| | BackColor | &H80000005& |
| | ImageHeight | 0 |
| | ImageWidth | 0 |
| | Tag | |
| **Posição** | Height | 28,5 |
| | Left | 606 |
| | Top | 108 |
| | Width | 28,5 |

### List View

> ListView: Controle avançado utilizado para exibir uma coleção de itens em quatro modos de visualização distintos: Ícones Grandes, Ícones Pequenos, Lista e Relatório (Report). No modo Relatório, permite a criação de tabelas complexas com múltiplas colunas, cabeçalhos clicáveis para ordenação e seleção de linha inteira (FullRowSelect). É amplamente utilizado em sistemas de gestão para listar registros de bancos de dados com suporte a ícones provenientes de um ImageList associado.

| Categoria | Propriedade | Valor |
| :--- | :--- | :--- |
| **Aparência** | (Name) | ListView1 |
| | ControlTipText | |
| | Visible | True |
| **Comportamento** | AllowColumnReorder | False |
| | Arrange | 0 - lvwNone |
| | FullRowSelect | False |
| | HideColumnHeaders | False |
| | HideSelection | True |
| | HotTracking | False |
| | HoverSelection | False |
| | LabelEdit | 0 - lvwAutomatic |
| | LabelWrap | True |
| | MultiSelect | False |
| | Sorted | False |
| | SortKey | 0 |
| | SortOrder | 0 - lvwAscending |
| **Diversos** | (Personalizado) | |
| | (Sobre) | |
| | Appearance | 1 - cc3D |
| | BackColor | &H80000005& |
| | BorderStyle | 1 - ccFixedSingle |
| | Checkboxes | False |
| | Enabled | True |
| | FlatScrollBar | False |
| | Font | Tahoma |
| | ForeColor | &H80000008& |
| | GridLines | False |
| | HelpContextID | 0 |
| | MouseIcon | (None) |
| | MousePointer | 0 - ccDefault |
| | OLEDDragMode | 0 - ccOLEDragManual |
| | OLEDDropMode | 0 - ccOLEDropNone |
| | Picture | (Nenhum) |
| | PictureAlignment | 0 - lvwTopLeft |
| | TabIndex | 23 |
| | TabStop | True |
| | Tag | |
| | TextBackground | 0 - lvwTransparent |
| | View | 0 - lvwIcon |
| **Posição** | Height | 37,5 |
| | Left | 600 |
| | Top | 138 |
| | Width | 75 |

### Progress Bar

> ProgressBar: Controlo visual utilizado para indicar o progresso de uma operação demorada através do preenchimento gradual de uma barra. Permite configurar valores mínimos e máximos (geralmente representando 0% e 100%) e pode ser exibido em orientações horizontal ou vertical. É essencial para fornecer feedback visual ao utilizador sobre o estado de processos como a importação de dados ou cálculos extensos no VBA.

| Categoria | Propriedade | Valor |
| :--- | :--- | :--- |
| **Aparência** | (Name) | ProgressBar1 |
| | ControlTipText | |
| | Orientation | 0 - ccOrientationHorizontal |
| | Scrolling | 0 - ccScrollingStandard |
| | Visible | True |
| **Diversos** | (Personalizado) | |
| | (Sobre) | |
| | Appearance | 1 - cc3D |
| | BorderStyle | 0 - ccNone |
| | Enabled | True |
| | HelpContextID | 0 |
| | Max | 100 |
| | Min | 0 |
| | MouseIcon | (None) |
| | MousePointer | 0 - ccDefault |
| | OLEDropMode | 0 - ccOLEDropNone |
| | TabIndex | 23 |
| | TabStop | False |
| | Tag | |
| **Posição** | Height | 37,5 |
| | Left | 648 |
| | Top | 96 |
| | Width | 75 |

### Slider

> Slider: Controle ActiveX que permite ao usuário selecionar um valor dentro de um intervalo contínuo movendo um botão deslizante ao longo de uma trilha. É ideal para ajustes de magnitude (como volume, intensidade ou zoom) e oferece suporte a marcas de escala (ticks) para orientação visual. Diferente da ScrollBar, o Slider é focado na seleção de um valor específico em um espectro, em vez de navegação de conteúdo.

| Categoria | Propriedade | Valor |
| :--- | :--- | :--- |
| **Aparência** | (Name) | Slider1 |
| | ControlTipText | |
| | Orientation | 0 - ccOrientationHorizontal |
| | TextPosition | 0 - sldAboveLeft |
| | TickFrequency | 1 |
| | TickStyle | 0 - sldBottomRight |
| | Visible | True |
| **Comportamento** | LargeChange | 5 |
| | SelectRange | False |
| | SelLength | 0 |
| | SelStart | 0 |
| | SmallChange | 1 |
| **Diversos** | (Personalizado) | |
| | (Sobre) | |
| | BorderStyle | 0 - ccNone |
| | Enabled | True |
| | HelpContextID | 0 |
| | Max | 10 |
| | Min | 0 |
| | MouseIcon | (None) |
| | MousePointer | 0 - ccDefault |
| | OLEDropMode | 0 - ccOLEDropNone |
| | TabIndex | 23 |
| | TabStop | True |
| | Tag | |
| | Value | 0 |
| **Posição** | Height | 37,5 |
| | Left | 636 |
| | Top | 132 |
| | Width | 75 |

### Status Bar

> StatusBar: Controle localizado geralmente na base de um formulário que exibe informações de status, mensagens de ajuda ou progresso de tarefas. Pode ser configurado como um painel único (SimpleText) ou dividido em múltiplos painéis para mostrar diferentes tipos de dados simultaneamente (como data, hora ou estado de teclas de bloqueio), sendo essencial para manter o usuário informado sobre o estado da aplicação sem interromper o fluxo de trabalho.

| Categoria | Propriedade | Valor |
| :--- | :--- | :--- |
| **Aparência** | (Name) | StatusBar1 |
| | ControlTipText | |
| | Visible | True |
| **Comportamento** | ShowTips | True |
| | Style | 0 - sbrNormal |
| **Diversos** | (Personalizado) | |
| | (Sobre) | |
| | Enabled | True |
| | Font | Tahoma |
| | HelpContextID | 0 |
| | MouseIcon | (None) |
| | MousePointer | 0 - ccDefault |
| | OLEDropMode | 0 - ccOLEDropNone |
| | SimpleText | |
| | TabIndex | 23 |
| | TabStop | False |
| | Tag | |
| **Posição** | Height | 18,75 |
| | Left | 582 |
| | Top | 126 |
| | Width | 75 |

### Tab Strip (Variação)

> TabStrip (Common Controls): Esta variação do controle TabStrip pertence à biblioteca Microsoft Windows Common Controls. Diferente da versão padrão do MS Forms, ela permite um controle mais refinado sobre o posicionamento das abas (propriedade Placement), suporte a separadores visuais e diferentes modos de dimensionamento das guias (TabWidthStyle). É ideal para criar interfaces que precisam seguir estritamente o visual nativo do sistema operacional Windows.

| Categoria | Propriedade | Valor |
| :--- | :--- | :--- |
| **Aparência** | (Name) | TabStrip1 |
| | ControlTipText | |
| | Placement | 0 - tabPlacementTop |
| | Separators | False |
| | Style | 0 - tabTabs |
| | Visible | True |
| **Comportamento** | HotTracking | False |
| | MultiSelect | False |
| | ShowTips | True |
| | TabStyle | 0 - tabTabStandard |
| **Diversos** | (Personalizado) | |
| | (Sobre) | |
| | Enabled | True |
| | Font | Tahoma |
| | HelpContextID | 0 |
| | MouseIcon | (None) |
| | MousePointer | 0 - ccDefault |
| | OLEDropMode | 0 - ccOLEDropNone |
| | TabIndex | 23 |
| | TabMinWidth | 28,35 |
| | TabStop | True |
| | Tag | |
| **Posição** | Height | 37,5 |
| | Left | 594 |
| | Top | 90 |
| | Width | 75 |
| **Tabs** | MultiRow | False |
| | TabFixedHeight | 0 |
| | TabFixedWidth | 0 |
| | TabWidthStyle | 0 - tabJustified |

### Toolbar

> Toolbar: Controle ActiveX que permite a criação de barras de ferramentas contendo uma coleção de objetos Button. É frequentemente associado a um controle ImageList para exibir ícones nos botões. Oferece funcionalidades avançadas como a criação de botões de alternância (toggle), grupos de botões (estilo rádio), menus suspensos embutidos e a capacidade de o usuário personalizar a barra em tempo de execução (AllowCustomize).

| Categoria | Propriedade | Valor |
| :--- | :--- | :--- |
| **Outros** | TextAlign | 0 - tbrTextAlignBottom |
| **Aparência** | (Name) | Toolbar1 |
| | ButtonHeight | 16,5005 |
| | ButtonWidth | 18,0005 |
| | ControlTipText | |
| | Style | 0 - tbrStandard |
| | Visible | True |
| **Comportamento** | AllowCustomize | True |
| | ShowTips | True |
| | Wrappable | True |
| **Diversos** | (Personalizado) | |
| | (Sobre) | |
| | Appearance | 1 - cc3D |
| | BorderStyle | 0 - ccNone |
| | Enabled | True |
| | HelpContextID | 0 |
| | HelpFile | |
| | MouseIcon | (None) |
| | MousePointer | 0 - ccDefault |
| | OLEDropMode | 0 - ccOLEDropNone |
| | TabIndex | 23 |
| | TabStop | False |
| | Tag | |
| **Posição** | Height | 37,5 |
| | Left | 570 |
| | Top | 120 |
| | Width | 75 |

### Tree View

> TreeView: Controle ActiveX utilizado para exibir uma representação hierárquica de itens, denominados "Nodes" (Nós). Cada nó pode conter texto e imagens (via ImageList) e pode ser expandido ou recolhido para revelar ou ocultar nós filhos. É a ferramenta padrão para criar navegadores de arquivos, organogramas ou sistemas de categorias aninhadas, oferecendo suporte a eventos complexos de seleção e drag-and-drop.

| Categoria | Propriedade | Valor |
| :--- | :--- | :--- |
| **Aparência** | (Name) | TreeView1 |
| | Checkboxes | False |
| | ControlTipText | |
| | HotTracking | False |
| | Indentation | 28,35 |
| | LineStyle | 0 - tvwTreeLines |
| | Visible | True |
| **Comportamento** | FullRowSelect | False |
| | HideSelection | True |
| | LabelEdit | 0 - tvwAutomatic |
| | Scroll | True |
| | SingleSel | False |
| | Sorted | False |
| | Style | 7 - tvwTreelinesPlusMinusPictureText |
| **Diversos** | (Personalizado) | |
| | (Sobre) | |
| | Appearance | 1 - cc3D |
| | BorderStyle | 0 - ccNone |
| | Enabled | True |
| | Font | Tahoma |
| | HelpContextID | 0 |
| | MouseIcon | (None) |
| | MousePointer | 0 - ccDefault |
| | OLEDDragMode | 0 - ccOLEDragManual |
| | OLEDropMode | 0 - ccOLEDropNone |
| | PathSeparator | \ |
| | TabIndex | 23 |
| | TabStop | True |
| | Tag | |
| **Posição** | Height | 138 |
| | Left | 636 |
| | Top | 90 |
| | Width | 132 |

## Considerações

> Relato de um áudio

{
> Como você pode observar anteriormente, tive que documentar vários detalhes do Excel, do projeto em si, etc. Eu quero consolidar tudo e deixar robusto. Porém, eu estou num impasse de utilizar apenas o Excel ou só o Access, ou os dois juntos de forma integrada, usando aí uma sincronização bidirecional através do Google Drive. Eu fico na dúvida o que fazer. Já que a empresa não vai lidar com milhares e milhares de dados, vamos supor, no máximo que nós chegamos é 250 alunos. Mas vamos supor que a média de tempo de curso que cada um faz é três anos. Isso são registros aí de cada aluno, não acho que não vai passar aí de 1.500 registros cada um, no total, seja financeiro, presença, falta. Geralmente é isso. Então não vai ser uma coisa escalável, vai ser só local, usada aí no máximo em dois computadores. Acredito que três, mas três vai ser raro. Então eu estou indeciso, entendeu? Pra fazer o esquema do banco de dados, o back-end do projeto. E outra coisa que eu preciso ver é criar interface amigável, como se eu tivesse desenvolvendo em React, por exemplo, usando Electron.js, que eu não vou usar nesse caso, eu vou usar apenas ferramentas puramente Office, né? Ou usar o Python também, se pudesse ajudar, mas não sei, eu quero evitar fazer gambiarra maluca ou armadilha negra pra funcionar o sistema, porque quanto mais puro ele melhor, acredito eu. Mas também eu aceito alguma biblioteca externa que pudesse um facilitador, um utilitário que pudesse ajudar ali no código VBA. que vai ser a maioria das coisas vai ser em VBA, não Python. Se tiver que usar o Python, seja em caso extremo, entendeu? Ou pegar alguma coisa de fora, web, mas eu acho que o VBA consegue dar conta de tudo, eu acredito eu, pra desenvolver tudo isso. Então, é, eu fico na dúvida do que fazer. Eu pensei em criar um ribbon, Excel ribbon customizado ou um Access ribbon, usando ali o Visual Studio direto, entendeu? Criar um projeto e manipular ali, não sei, mas já estou escalando um projeto, não escalando de tamanho de usuários, mas sim já estou aumentando a complexidade, né? Talvez seja até mais eficiente eu fazer isso. Criar uma ribbon customizada, né? Acho que seria interessante. né? É mais ou menos, eu não sei o que fazer, como é que eu vou fazer a interface. Porque Excel não foi feito pra criar programas, né? Ele foi feito pra você manipular dados e limpar dados. Mas pode até criar programas. Tem muita gente que faz projetos inteiros e vende softwares, na verdade sistemas, né, criados juntamente ali, puramente em Excel e Access. E tem gente que vende e ganha dinheiro com isso. Mas eu não sou especialista nisso ainda, apesar de eu ter um conhecimento legal em Excel, mas eu tenho um pouco em Access, mas eu sei ali mais ou menos SQL, e eu não sou especialista em VBA, eu sou iniciante ainda, né? Então eu preciso ver o que que é mais eficiente para o meu caso de uso, considerando todas as nuances que vimos até agora, né? Eu não sei como é que eu vou decidir tudo isso ainda. Abaixo eu vou deixar um link também pra você ver onde eu vejo questão de componentes, né, como usar o Excel em algumas questões também, o Excel e componentes do Excel e etc, VBA. para que você possa investigar aprofundadamente, ok? Ah, e também, abaixo, você pode ver duas tabelas em HTML, que vão ser os templates pra imprimir, como a gente falou lá anteriormente, né? Que vai imprimir na folha A4 de forma horizontal, para que essa tabela aí de em forma de lista, né? Que cada linha é o nome de aluno pra marcar a presença dele, dependendo do dia, porque ele vai servir tanto pra preencher no Excel, ou no Excel, eu não sei se dá pra fazer isso também, tipo abrir Excel dentro do Excel, não sei se é possível, né? Ou tipo, a pessoa vai preencher no Excel pra deixar um tipo um, eu não sei, fazer uma interface robusta e rígida, pra não ficar manipulando nem customizando nem nada, entendeu? É o que é. É como se fosse um front-end legal ali, né? Ah, e outra coisa, pra editar componentes de UI no Excel é difícil, porque eles são muito, como é que dizem, engessados, entendeu? Não sei se dá pra customizar ele pra deixar um pouco mais bonito. Só que eu jito que boniteza não tem necessidade, né? Pelo menos o mínimo ali de customização é necessário. Acredito que já seja o suficiente. Né, porque eu já estou fazendo outro projeto muito grande usando Electron.js, o banco de dados SQLite, que é paralelo a esse, mas eu preciso fazer uma solução rápida até sexta que vem, que é usando o Excel, o Excel. De preferência, apenas o Excel, mas eu eu não sei se eu de fato uso o Excel ou não. Entendeu? Então essa é a minha dúvida. Acabei de criar o arquivo BD_wiz_admin.xlsx limpo e sem nada. Sem tabela, planilhas e sem vba. Bom, para te dar clarificações necessárias, em questão de sincronização, eu vou usar mais o OneDrive, né? Porque aí eu vou sincronizar alguns documentos da empresa, né? Que vai usar em dois notebooks e no PC de mesa na recepção, entendeu? Na verdade, um notebook aqui eu vou usar mais, o outro é mais pela assessoria comercial, então não vai fazer diferença, mas será mais assim, no notebook 1 e no PC 1. O notebook 2, acho que aí vai ser difícil eu usar lá, só se for de extrema necessidade, mas a princípio vai ser apenas dois dispositivos. É porque assim, na verdade, a conta do Microsoft 365 é minha, só que eu compartilhei com uma pequena empresa, né? É uma empresa pequena, é só o e-mail ali, né? Aí eles podem usar, né? A minha equipe pode usar também ali a conta da empresa pra usar o Microsoft 365 Family, entendeu? A ficha de frequência impressa, ela vai continuar imprimindo sim, em papel e marcando a mão. É só pra, é só pra ter uma certeza. Tipo assim, se tiver algum erro, et cetera, a gente pode fazer umachecagem e ver se tá tudo certo. Ela vai registrar direto no Excel, entendeu? E a ficha de frequência é a prioridade não absoluta máxima, mas ela vai ser de alta prioridade. Entendeu? O gerador de ficha de frequência não tem a ver com isso, é mais a tabela que eu te mandei em HTML, os dois exemplos, por exemplo. Na verdade, assim, o que eu quero fazer é que, por exemplo, esses dois HTMLs, um é na verdade uma tabela pura, né? É normal. E outro é uma dinâmica que pega dessa. Entendeu? Então, baseando as configurações de cada aluno, por exemplo, vai ter a tabelinha para os horários de manhã, cada tabela vai ser uma hora. Tabela das 7 às 8. Na verdade, eu vou fazer diferente. Cada coluna, você pode ver ali que depois tem a questão dos dias do mês, né? Pode ser do dia 1 até o dia 31. Se for um mês de 28 dias, como fevereiro, vai exibir só 28 dias. Então, por exemplo, vai ser a folha no formato horizontal, A4, com a borda de 0,0, na verdade, 0,3 centímetros, né? Que eu vou declarar ali para padrão. E em cima, no cabeçalho, vai ter ali janeiro e o horário. Tipo assim, aí vai ter, como se fosse na segunda, vai ter as barrinhas ali, né? Vai ser as filas, né? Por exemplo, das 7 às 8, depois das 8 às 9, 9 às 10, 10 às 11. Aí parou. Aí no lado verso da folha, né? Eu vou colocar a tarde e continuar até, se não caber, vai continuar na página 3, entendeu? Então, eu quero automatizar essa criação com base na configuração de horários dos alunos, pra ser de forma muito rápida, entendeu? Essa é a questão. Porque eu estou pensando em usar, criar mais outro arquivo do Excel. Um é o pedagógico que eu uso e um é o na recepção. Os dois vão usar a mesma fonte, entendeu? Então, não vai ser o mesmo arquivo que vai ser compartilhado, que vai ser usado por duas máquinas. A pasta de trabalho de Excel, as duas vão ser com macros, né? XLSM, a minha que eu vou usar no notebook e a outra que vai ser usada na recepção. no computador, no PC de mesa, entendeu? Vão ser dois projetos que têm semelhanças, mas eles têm propósitos um pouco muito parecidos. Mas a questão de quem vai usar. Então, para evitar esse problema de que um arquivo está aberto através do OneDrive por duas máquinas, é a solução que eu tenho pra evitar essa edição simultânea. Entendeu? Então... e eles vão compartilhar a mesma pasta de Excel que eu coloquei o nome abaixo pra você, entendeu? Ah, sobre a questão do ElectronJS paralelo, isso aí, na verdade, eu vou só fazer quando eu terminar essa questão do Excel, entendeu? Esse mini projeto que eu estou falando pra você. E também sobre os idiomas, relatórios, se ele precisa ter ou não separar por idioma, na verdade, não tem necessidade, entendeu? Eu não vou querer gerar relatórios. Esses relatórios que eu falei pra você de... na verdade, pode até gerar, entendeu? Mas eu acho que não tem necessidade porque na verdade, se eu visualizar eles, apenas já ser suficiente contabilizar numa planilha aí de relatórios, entendeu? Eu acho que Power BI não precisa usar. Mas então vai ser uma coisa mais visual do que impressa, do que impresso, entendeu? É só a ficha da recepção mesmo, tá? E abaixo, eu vou te falar o nome do arquivo, né? Que é este. E aí eu quero que você me dê agora o passo a passo, né? Pra criar tudo assim, de forma rápida. As tabelas, a relação. E eu traduzo aqui o que eu tenho, que é o seguinte, o Excel, você sabe que ele tem um Power Pivot, onde eu organizo os dados dele ali. E eu posso manipular as tabelas, por exemplo, uma tabela ali, que eu tipo assim, eu crio uma tabela e eu posso formatar como tabela, né? Aí eu posso transformar essa tabela em um modelo de dados que fica no Power Pivot dentro do Excel. Aí, ele permite igual no Excel eu fazer relações, né? Como se fosse um banco de dados. Por exemplo, eu pego uma coluna ali, né? É como se fosse um diagrama de classe, entendeu? É um é é aquela diagrama de entidade relacional. É é literalmente isso. Por exemplo, pega ali um atributo, né? Que é uma coluna e arrasto para outro. Então o Excel permite fazer isso. Eu quero saber se na verdade eu preciso fazer isso tendo o VBA, que faz isso pra mim, ou é bom fazer as dois coisas pra não me confundir. Que que você acha melhor? E também eu quero que você me dê o VBA para criar tudo de uma vez, as planilhas necessárias com os nomes que você já viu no no no no meu relato, né? No arquivo de documento que eu te mandei e também é o nome das tabelas que eu também coloquei, já padronizei tudo, né? E se você quiser criar algo além disso também. Ah, e também a e já considerar o modelo de dados dessas tabelas e fazer as relações delas, né? Eh da forma robusta, entendeu? É da forma certa pra evitar aí duplicação de dados e e haver e conseguir haver integridade de dados. Tá, então, eh veja a parte um, que é tipo assim o o sealing, né? Ou criação do schema tudo de uma vez, se possível, né? Não sei se dá pra fazer com VBA, se você quiser que eu faça na mão, eu faço na mão mesmo. 

> As planilhas precisam ter o termo "BD_" antes do nome e o header de cada tabela precisa começar na linha 1

}

### Links úteis

- https://learn.microsoft.com/en-us/office/client-developer/access/access-home
- https://techcommunity.microsoft.com/blog/accessblog/try-the-new-modern-sql-editor-in-access/4289430
- https://learn.microsoft.com/en-us/office/vba/api/overview/excel
- https://learn.microsoft.com/en-us/office/vba/api/overview/excel/object-model
- https://support.microsoft.com/en-us/office/create-a-chart-on-a-form-or-report-1a463106-65d0-4dbb-9d66-4ecb737ea7f7
- https://learn.microsoft.com/en-us/office/vba/api/access.form
- https://learn.microsoft.com/en-us/office/vba/api/access.form.controls
- https://learn.microsoft.com/en-us/office/vba/api/overview/activex-control
- https://www.gigasoft.com/chart-activex-access
- https://support.microsoft.com/en-us/office/overview-of-forms-form-controls-and-activex-controls-on-a-worksheet-15ba7e28-8d7f-42ab-9470-ffb9ab94e7c2
- https://nolongerset.com/create-activex-control-with-twinbasic/

### Tabela 1

```html
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Ficha de Frequência Recepção - Janeiro 2026</title>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }

  body {
    font-family: Calibri, "Segoe UI", Arial, sans-serif;
    font-size: 11px;
    color: #000;
    background: #fff;
    padding: 10px;
  }

  .container {
    overflow-x: auto;
  }

  table {
    border-collapse: collapse;
    white-space: nowrap;
  }

  td, th {
    border: 1px solid #b8b8b8;
    padding: 2px 5px;
    vertical-align: middle;
  }

  /* ── Title row ── */
  .title-row td {
    border: none;
    font-weight: bold;
    padding: 4px 5px;
  }
  .title-text {
    font-size: 16px;
  }
  .title-month {
    font-size: 18px;
    text-align: right;
    font-style: italic;
    letter-spacing: 1px;
  }
  .title-year {
    font-size: 26px;
    text-align: right;
    font-weight: 900;
  }

  /* ── Blank spacer rows ── */
  .spacer-row td {
    border: none;
    height: 6px;
  }

  /* ── Date reference row ── */
  .date-row td {
    border: none;
    font-size: 10px;
    color: #333;
    padding: 1px 5px;
  }

  /* ── Day-of-week row ── */
  .dow-row td {
    background: #fff2cc;
    text-align: center;
    font-size: 10px;
    font-weight: bold;
    padding: 2px 1px;
    color: #333;
  }
  .dow-row td.empty-dow {
    background: transparent;
    border: none;
  }
  .dow-row td.dom {
    background: #fde9d9;
    color: #c00;
  }

  /* ── Header row ── */
  .header-row th {
    background: #fce4b5;
    font-weight: bold;
    text-align: center;
    padding: 3px 5px;
    font-size: 11px;
  }
  .header-row th.day-num {
    min-width: 22px;
    width: 22px;
    padding: 3px 1px;
    font-size: 10px;
  }
  .header-row th.dom-num {
    background: #f5d0a9;
    color: #c00;
  }

  /* ── Data rows ── */
  .data-row td {
    padding: 2px 5px;
    height: 20px;
  }
  .data-row td.nome {
    text-align: left;
    min-width: 160px;
  }
  .data-row td.center {
    text-align: center;
  }
  .data-row td.status {
    text-align: center;
    min-width: 40px;
  }
  .data-row td.estagio {
    text-align: left;
    min-width: 105px;
  }
  .data-row td.modalidade {
    text-align: left;
    min-width: 80px;
  }
  .data-row td.professor {
    text-align: left;
    min-width: 60px;
  }
  .data-row td.dias {
    text-align: center;
    min-width: 60px;
    font-size: 10px;
  }
  .data-row td.check {
    text-align: center;
    min-width: 28px;
    width: 28px;
  }
  .data-row td.hora {
    text-align: center;
    min-width: 38px;
    font-weight: bold;
  }
  .data-row td.day-cell {
    text-align: center;
    min-width: 22px;
    width: 22px;
    padding: 2px 1px;
  }
  .data-row td.dom-cell {
    background: #fdf5ed;
  }
</style>
</head>
<body>
<div class="container">
<table>
  <!-- ══════ TITLE ══════ -->
  <tr class="title-row">
    <td class="title-text" colspan="13">Ficha de Frequência Recepção</td>
    <td class="title-month" colspan="24">Janeiro</td>
    <td class="title-year" colspan="7">2026</td>
  </tr>

  <!-- spacer -->
  <tr class="spacer-row"><td colspan="44"></td></tr>

  <!-- ══════ DATE REF ══════ -->
  <tr class="date-row">
    <td colspan="3"></td>
    <td>30/01/2026</td>
    <td>30/01/2026</td>
    <td colspan="39"></td>
  </tr>

  <!-- spacer -->
  <tr class="spacer-row"><td colspan="44"></td></tr>

  <!-- ══════ DAY-OF-WEEK ROW ══════ -->
  <tr class="dow-row">
    <td class="empty-dow" colspan="13"></td>
    <td>qui</td><!--1-->
    <td>sex</td><!--2-->
    <td>sáb</td><!--3-->
    <td class="dom">dom</td><!--4-->
    <td>seg</td><!--5-->
    <td>ter</td><!--6-->
    <td>qua</td><!--7-->
    <td>qui</td><!--8-->
    <td>sex</td><!--9-->
    <td>sáb</td><!--10-->
    <td class="dom">dom</td><!--11-->
    <td>seg</td><!--12-->
    <td>ter</td><!--13-->
    <td>qua</td><!--14-->
    <td>qui</td><!--15-->
    <td>sex</td><!--16-->
    <td>sáb</td><!--17-->
    <td class="dom">dom</td><!--18-->
    <td>seg</td><!--19-->
    <td>ter</td><!--20-->
    <td>qua</td><!--21-->
    <td>qui</td><!--22-->
    <td>sex</td><!--23-->
    <td>sáb</td><!--24-->
    <td class="dom">dom</td><!--25-->
    <td>seg</td><!--26-->
    <td>ter</td><!--27-->
    <td>qua</td><!--28-->
    <td>qui</td><!--29-->
    <td>sex</td><!--30-->
    <td>sáb</td><!--31-->
  </tr>

  <!-- ══════ HEADER ROW ══════ -->
  <tr class="header-row">
    <th>Aluno(a)</th>
    <th>Status</th>
    <th>Estágio</th>
    <th>Modalidade</th>
    <th>Professor</th>
    <th>Dias</th>
    <th>Seg</th>
    <th>Ter</th>
    <th>Qua</th>
    <th>Qui</th>
    <th>Sex</th>
    <th>Sáb</th>
    <th>Hora</th>
    <th class="day-num">1</th>
    <th class="day-num">2</th>
    <th class="day-num">3</th>
    <th class="day-num dom-num">4</th>
    <th class="day-num">5</th>
    <th class="day-num">6</th>
    <th class="day-num">7</th>
    <th class="day-num">8</th>
    <th class="day-num">9</th>
    <th class="day-num">10</th>
    <th class="day-num dom-num">11</th>
    <th class="day-num">12</th>
    <th class="day-num">13</th>
    <th class="day-num">14</th>
    <th class="day-num">15</th>
    <th class="day-num">16</th>
    <th class="day-num">17</th>
    <th class="day-num dom-num">18</th>
    <th class="day-num">19</th>
    <th class="day-num">20</th>
    <th class="day-num">21</th>
    <th class="day-num">22</th>
    <th class="day-num">23</th>
    <th class="day-num">24</th>
    <th class="day-num dom-num">25</th>
    <th class="day-num">26</th>
    <th class="day-num">27</th>
    <th class="day-num">28</th>
    <th class="day-num">29</th>
    <th class="day-num">30</th>
    <th class="day-num">31</th>
  </tr>

  <!-- ══════ DATA ROWS ══════ -->

  <!-- Adelita -->
  <tr class="data-row">
    <td class="nome">Adelita</td>
    <td class="status">T</td>
    <td class="estagio">W8</td>
    <td class="modalidade">Interactive</td>
    <td class="professor">Vitor</td>
    <td class="dias">2ª | Sáb</td>
    <td class="check">x</td>
    <td class="check"></td>
    <td class="check"></td>
    <td class="check"></td>
    <td class="check"></td>
    <td class="check">x</td>
    <td class="hora">08:00</td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

  <!-- Alencássia Peres -->
  <tr class="data-row">
    <td class="nome">Alencássia Peres</td>
    <td class="status">T</td>
    <td class="estagio">W2 New</td>
    <td class="modalidade">Interactive</td>
    <td class="professor">Vitor</td>
    <td class="dias"></td>
    <td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td>
    <td class="hora">08:00</td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

  <!-- Aline Lucena -->
  <tr class="data-row">
    <td class="nome">Aline Lucena</td>
    <td class="status">C</td>
    <td class="estagio">ESP4</td>
    <td class="modalidade">Interactive</td>
    <td class="professor">Vitor</td>
    <td class="dias"></td>
    <td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td>
    <td class="hora">08:00</td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

  <!-- Ana Andrade -->
  <tr class="data-row">
    <td class="nome">Ana Andrade</td>
    <td class="status">C</td>
    <td class="estagio">TOTS 2</td>
    <td class="modalidade">Interactive</td>
    <td class="professor">Vitor</td>
    <td class="dias"></td>
    <td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td>
    <td class="hora">08:00</td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

  <!-- Ana Biatriz -->
  <tr class="data-row">
    <td class="nome">Ana Biatriz</td>
    <td class="status">A</td>
    <td class="estagio">NEXT GEN.</td>
    <td class="modalidade">Interactive</td>
    <td class="professor">Vitor</td>
    <td class="dias">2ª | 4ª</td>
    <td class="check">x</td>
    <td class="check"></td>
    <td class="check">x</td>
    <td class="check"></td>
    <td class="check"></td>
    <td class="check"></td>
    <td class="hora">15:00</td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

  <!-- Ana Diniz -->
  <tr class="data-row">
    <td class="nome">Ana Diniz</td>
    <td class="status">C</td>
    <td class="estagio">W4 Old</td>
    <td class="modalidade">Interactive</td>
    <td class="professor">Vitor</td>
    <td class="dias"></td>
    <td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td>
    <td class="hora">08:00</td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

  <!-- Ana Joyce -->
  <tr class="data-row">
    <td class="nome">Ana Joyce</td>
    <td class="status">T</td>
    <td class="estagio">W8</td>
    <td class="modalidade">Interactive</td>
    <td class="professor">Vitor</td>
    <td class="dias"></td>
    <td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td>
    <td class="hora">08:00</td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

  <!-- Ana Maria -->
  <tr class="data-row">
    <td class="nome">Ana Maria</td>
    <td class="status">C</td>
    <td class="estagio">W2 New</td>
    <td class="modalidade">Interactive</td>
    <td class="professor">Vitor</td>
    <td class="dias"></td>
    <td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td>
    <td class="hora">08:00</td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

  <!-- Ana Paula -->
  <tr class="data-row">
    <td class="nome">Ana Paula</td>
    <td class="status">C</td>
    <td class="estagio">W2 New</td>
    <td class="modalidade">Interactive</td>
    <td class="professor">Vitor</td>
    <td class="dias"></td>
    <td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td>
    <td class="hora">08:00</td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

  <!-- André Roguigues -->
  <tr class="data-row">
    <td class="nome">André Roguigues</td>
    <td class="status">T</td>
    <td class="estagio">W2 New</td>
    <td class="modalidade">Interactive</td>
    <td class="professor">Vitor</td>
    <td class="dias"></td>
    <td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td>
    <td class="hora">08:00</td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

  <!-- Angelina Figueiredo -->
  <tr class="data-row">
    <td class="nome">Angelina Figueiredo</td>
    <td class="status">A</td>
    <td class="estagio">TEENS 2 New</td>
    <td class="modalidade">Interactive</td>
    <td class="professor">Vitor</td>
    <td class="dias">3ª | 5ª</td>
    <td class="check"></td>
    <td class="check">x</td>
    <td class="check"></td>
    <td class="check">x</td>
    <td class="check"></td>
    <td class="check"></td>
    <td class="hora">15:00</td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

  <!-- Anna Clara -->
  <tr class="data-row">
    <td class="nome">Anna Clara</td>
    <td class="status">T</td>
    <td class="estagio">NEXT GEN.</td>
    <td class="modalidade">Interactive</td>
    <td class="professor">Vitor</td>
    <td class="dias"></td>
    <td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td>
    <td class="hora">08:00</td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

  <!-- Anne Stobienia -->
  <tr class="data-row">
    <td class="nome">Anne Stobienia</td>
    <td class="status">A</td>
    <td class="estagio">ESP2 Nuevo</td>
    <td class="modalidade">Interactive</td>
    <td class="professor">Vitor</td>
    <td class="dias">3ª | 5ª</td>
    <td class="check"></td>
    <td class="check">x</td>
    <td class="check"></td>
    <td class="check">x</td>
    <td class="check"></td>
    <td class="check"></td>
    <td class="hora">17:00</td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

  <!-- Antônio Moraes -->
  <tr class="data-row">
    <td class="nome">Antônio Moraes</td>
    <td class="status">A</td>
    <td class="estagio">W2 New</td>
    <td class="modalidade">Interactive</td>
    <td class="professor"></td>
    <td class="dias">2ª | 4ª</td>
    <td class="check">x</td>
    <td class="check"></td>
    <td class="check">x</td>
    <td class="check"></td>
    <td class="check"></td>
    <td class="check"></td>
    <td class="hora">17:00</td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

  <!-- Antônio Neto -->
  <tr class="data-row">
    <td class="nome">Antônio Neto</td>
    <td class="status">A</td>
    <td class="estagio">NEXT GEN.</td>
    <td class="modalidade">Interactive</td>
    <td class="professor"></td>
    <td class="dias">4ª | 6ª</td>
    <td class="check"></td>
    <td class="check"></td>
    <td class="check">x</td>
    <td class="check"></td>
    <td class="check">x</td>
    <td class="check"></td>
    <td class="hora">13:00</td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

  <!-- Arthur Flores -->
  <tr class="data-row">
    <td class="nome">Arthur Flores</td>
    <td class="status">A</td>
    <td class="estagio">TEENS 4 New</td>
    <td class="modalidade">Interactive</td>
    <td class="professor"></td>
    <td class="dias">3ª | 5ª</td>
    <td class="check"></td>
    <td class="check">x</td>
    <td class="check"></td>
    <td class="check">x</td>
    <td class="check"></td>
    <td class="check"></td>
    <td class="hora">14:00</td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

  <!-- Ayllon -->
  <tr class="data-row">
    <td class="nome">Ayllon</td>
    <td class="status">A</td>
    <td class="estagio">W6</td>
    <td class="modalidade">Interactive</td>
    <td class="professor"></td>
    <td class="dias">2ª | Sáb</td>
    <td class="check">x</td>
    <td class="check"></td>
    <td class="check"></td>
    <td class="check"></td>
    <td class="check"></td>
    <td class="check">x</td>
    <td class="hora">13:00</td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

  <!-- Ayra Miamoto -->
  <tr class="data-row">
    <td class="nome">Ayra Miamoto</td>
    <td class="status">A</td>
    <td class="estagio">ESP2 Nuevo</td>
    <td class="modalidade">Interactive</td>
    <td class="professor"></td>
    <td class="dias">2ª | 3ª</td>
    <td class="check">x</td>
    <td class="check">x</td>
    <td class="check">x</td>
    <td class="check"></td>
    <td class="check"></td>
    <td class="check"></td>
    <td class="hora">07:00</td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

  <!-- Barbara Assis -->
  <tr class="data-row">
    <td class="nome">Barbara Assis</td>
    <td class="status">T</td>
    <td class="estagio">NEXT GEN.</td>
    <td class="modalidade">Interactive</td>
    <td class="professor"></td>
    <td class="dias"></td>
    <td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td>
    <td class="hora"></td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

  <!-- Barbara Barbosa -->
  <tr class="data-row">
    <td class="nome">Barbara Barbosa</td>
    <td class="status">C</td>
    <td class="estagio">NEXT GEN.</td>
    <td class="modalidade">Interactive</td>
    <td class="professor"></td>
    <td class="dias"></td>
    <td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td>
    <td class="hora"></td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

  <!-- Beatriz Oliveira -->
  <tr class="data-row">
    <td class="nome">Beatriz Oliveira</td>
    <td class="status">C</td>
    <td class="estagio">W2 Old</td>
    <td class="modalidade">Interactive</td>
    <td class="professor"></td>
    <td class="dias"></td>
    <td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td>
    <td class="hora"></td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

  <!-- Bernardo Becker -->
  <tr class="data-row">
    <td class="nome">Bernardo Becker</td>
    <td class="status">A</td>
    <td class="estagio">KIDS 2 New</td>
    <td class="modalidade">Interactive</td>
    <td class="professor"></td>
    <td class="dias">3ª | 5ª</td>
    <td class="check"></td>
    <td class="check">x</td>
    <td class="check"></td>
    <td class="check">x</td>
    <td class="check"></td>
    <td class="check"></td>
    <td class="hora">10:00</td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

  <!-- Bernardo Yamo -->
  <tr class="data-row">
    <td class="nome">Bernardo Yamo</td>
    <td class="status">T</td>
    <td class="estagio">TEENS 2 New</td>
    <td class="modalidade">Interactive</td>
    <td class="professor"></td>
    <td class="dias"></td>
    <td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td>
    <td class="hora"></td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

  <!-- Bianca Nakassugui -->
  <tr class="data-row">
    <td class="nome">Bianca Nakassugui</td>
    <td class="status">A</td>
    <td class="estagio">TEENS 6 New</td>
    <td class="modalidade">Interactive</td>
    <td class="professor"></td>
    <td class="dias">2ª | 6ª</td>
    <td class="check">x</td>
    <td class="check"></td>
    <td class="check"></td>
    <td class="check"></td>
    <td class="check">x</td>
    <td class="check"></td>
    <td class="hora">13:00</td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

  <!-- Bruna Nayara -->
  <tr class="data-row">
    <td class="nome">Bruna Nayara</td>
    <td class="status">A</td>
    <td class="estagio">W2 New</td>
    <td class="modalidade">Interactive</td>
    <td class="professor"></td>
    <td class="dias">6ª | Sáb</td>
    <td class="check"></td>
    <td class="check"></td>
    <td class="check"></td>
    <td class="check"></td>
    <td class="check">x</td>
    <td class="check">x</td>
    <td class="hora">09:00</td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

  <!-- Bryan Diniz -->
  <tr class="data-row">
    <td class="nome">Bryan Diniz</td>
    <td class="status">C</td>
    <td class="estagio">TOTS 2</td>
    <td class="modalidade">Interactive</td>
    <td class="professor"></td>
    <td class="dias"></td>
    <td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td><td class="check"></td>
    <td class="hora"></td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

</table>
</div>
</body>
</html>
```

### Tabela 2

```html
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Ficha de Frequência Recepção - Tabela Dinâmica - Janeiro 2026</title>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }

  body {
    font-family: Calibri, "Segoe UI", Arial, sans-serif;
    font-size: 11px;
    color: #000;
    background: #fff;
    padding: 10px;
  }

  .container {
    overflow-x: auto;
  }

  /* ══════ SLICERS ══════ */
  .slicers {
    display: flex;
    gap: 16px;
    flex-wrap: wrap;
    margin-bottom: 14px;
    align-items: flex-start;
  }

  .slicer {
    border: 1px solid #b8b8b8;
    border-radius: 2px;
    background: #fff;
    min-width: 90px;
  }

  .slicer-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 3px 6px;
    font-size: 11px;
    font-weight: bold;
    border-bottom: 1px solid #d0d0d0;
    background: #f5f5f5;
  }

  .slicer-header span {
    flex: 1;
  }

  .slicer-icons {
    display: flex;
    gap: 4px;
    margin-left: 8px;
  }

  .slicer-icon {
    font-size: 10px;
    color: #666;
    cursor: default;
  }

  .slicer-buttons {
    display: flex;
    padding: 3px 4px;
    gap: 3px;
  }

  .slicer-btn {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    padding: 2px 8px;
    font-size: 10px;
    border: 1px solid #b0b0b0;
    border-radius: 2px;
    background: #fff;
    color: #333;
    cursor: default;
    min-width: 28px;
    white-space: nowrap;
  }

  .slicer-btn.selected {
    background: #4472c4;
    color: #fff;
    border-color: #3a62a8;
  }

  .slicer-btn.partial {
    background: #b4c7e7;
    color: #1a1a1a;
    border-color: #8faadb;
  }

  /* ══════ TABLE ══════ */
  table {
    border-collapse: collapse;
    white-space: nowrap;
  }

  td, th {
    border: 1px solid #b8b8b8;
    padding: 2px 5px;
    vertical-align: middle;
  }

  /* ── Title row ── */
  .title-row td {
    border: none;
    font-weight: bold;
    padding: 4px 5px;
  }

  .title-text {
    font-size: 18px;
    font-style: italic;
  }

  .title-month {
    font-size: 22px;
    text-align: right;
    font-weight: 900;
  }

  /* ── Spacer ── */
  .spacer-row td {
    border: none;
    height: 6px;
  }

  /* ── Day-of-week row ── */
  .dow-row td {
    background: #fff2cc;
    text-align: center;
    font-size: 10px;
    font-weight: bold;
    padding: 2px 1px;
    color: #333;
  }
  .dow-row td.empty-dow {
    background: transparent;
    border: none;
  }
  .dow-row td.dom {
    background: #fde9d9;
    color: #c00;
  }

  /* ── Header row ── */
  .header-row th {
    background: #fce4b5;
    font-weight: bold;
    text-align: center;
    padding: 3px 5px;
    font-size: 11px;
  }
  .header-row th.day-num {
    min-width: 22px;
    width: 22px;
    padding: 3px 1px;
    font-size: 10px;
  }
  .header-row th.dom-num {
    background: #f5d0a9;
    color: #c00;
  }
  .header-row th .filter-icon {
    font-size: 8px;
    color: #666;
    margin-left: 2px;
  }

  /* ── Group header rows (hora) ── */
  .group-row td {
    background: #e2e2e2;
    font-weight: bold;
    font-size: 11px;
    padding: 3px 5px;
    border: 1px solid #b8b8b8;
  }

  /* ── Data rows ── */
  .data-row td {
    padding: 2px 5px;
    height: 20px;
  }
  .data-row td.nome {
    text-align: left;
    min-width: 150px;
  }
  .data-row td.estagio {
    text-align: center;
    min-width: 90px;
  }
  .data-row td.modalidade {
    text-align: left;
    min-width: 80px;
  }
  .data-row td.dias {
    text-align: center;
    min-width: 80px;
    font-size: 10px;
  }
  .data-row td.day-cell {
    text-align: center;
    min-width: 22px;
    width: 22px;
    padding: 2px 1px;
  }
  .data-row td.dom-cell {
    background: #fdf5ed;
  }
</style>
</head>
<body>
<div class="container">

<!-- ══════ SLICERS ══════ -->
<div class="slicers">
  <div class="slicer">
    <div class="slicer-header">
      <span>Seg</span>
      <div class="slicer-icons"><span class="slicer-icon">⇅</span><span class="slicer-icon">▽</span></div>
    </div>
    <div class="slicer-buttons">
      <span class="slicer-btn partial">(v...</span>
      <span class="slicer-btn selected">x</span>
    </div>
  </div>

  <div class="slicer">
    <div class="slicer-header">
      <span>Ter</span>
      <div class="slicer-icons"><span class="slicer-icon">⇅</span><span class="slicer-icon">▽</span></div>
    </div>
    <div class="slicer-buttons">
      <span class="slicer-btn partial">(v...</span>
      <span class="slicer-btn selected">x</span>
    </div>
  </div>

  <div class="slicer">
    <div class="slicer-header">
      <span>Qua</span>
      <div class="slicer-icons"><span class="slicer-icon">⇅</span><span class="slicer-icon">▽</span></div>
    </div>
    <div class="slicer-buttons">
      <span class="slicer-btn partial">(v...</span>
      <span class="slicer-btn selected">x</span>
    </div>
  </div>

  <div class="slicer">
    <div class="slicer-header">
      <span>Qui</span>
      <div class="slicer-icons"><span class="slicer-icon">⇅</span><span class="slicer-icon">▽</span></div>
    </div>
    <div class="slicer-buttons">
      <span class="slicer-btn partial">(v...</span>
      <span class="slicer-btn selected">x</span>
    </div>
  </div>

  <div class="slicer">
    <div class="slicer-header">
      <span>Sex</span>
      <div class="slicer-icons"><span class="slicer-icon">⇅</span><span class="slicer-icon">▽</span></div>
    </div>
    <div class="slicer-buttons">
      <span class="slicer-btn partial">(v...</span>
      <span class="slicer-btn selected">x</span>
    </div>
  </div>

  <div class="slicer">
    <div class="slicer-header">
      <span>Sáb</span>
      <div class="slicer-icons"><span class="slicer-icon">⇅</span><span class="slicer-icon">▽</span></div>
    </div>
    <div class="slicer-buttons">
      <span class="slicer-btn partial">(v...</span>
      <span class="slicer-btn">x</span>
    </div>
  </div>

  <div class="slicer" style="min-width: 120px;">
    <div class="slicer-header">
      <span>Status</span>
      <div class="slicer-icons"><span class="slicer-icon">⇅</span><span class="slicer-icon">▽</span></div>
    </div>
    <div class="slicer-buttons">
      <span class="slicer-btn selected">A</span>
      <span class="slicer-btn">C</span>
      <span class="slicer-btn">T</span>
    </div>
  </div>
</div>

<!-- ══════ TABLE ══════ -->
<table>
  <!-- Title -->
  <tr class="title-row">
    <td class="title-text" colspan="4">Segundas / Quartas</td>
    <td class="title-month" colspan="31">Janeiro</td>
  </tr>

  <!-- Spacer -->
  <tr class="spacer-row"><td colspan="35"></td></tr>

  <!-- Day-of-week row -->
  <tr class="dow-row">
    <td class="empty-dow" colspan="4"></td>
    <td>qui</td><!--1-->
    <td>sex</td><!--2-->
    <td>sáb</td><!--3-->
    <td class="dom">dom</td><!--4-->
    <td>seg</td><!--5-->
    <td>ter</td><!--6-->
    <td>qua</td><!--7-->
    <td>qui</td><!--8-->
    <td>sex</td><!--9-->
    <td>sáb</td><!--10-->
    <td class="dom">dom</td><!--11-->
    <td>seg</td><!--12-->
    <td>ter</td><!--13-->
    <td>qua</td><!--14-->
    <td>qui</td><!--15-->
    <td>sex</td><!--16-->
    <td>sáb</td><!--17-->
    <td class="dom">dom</td><!--18-->
    <td>seg</td><!--19-->
    <td>ter</td><!--20-->
    <td>qua</td><!--21-->
    <td>qui</td><!--22-->
    <td>sex</td><!--23-->
    <td>sáb</td><!--24-->
    <td class="dom">dom</td><!--25-->
    <td>seg</td><!--26-->
    <td>ter</td><!--27-->
    <td>qua</td><!--28-->
    <td>qui</td><!--29-->
    <td>sex</td><!--30-->
    <td>sáb</td><!--31-->
  </tr>

  <!-- Header row -->
  <tr class="header-row">
    <th>Alunos <span class="filter-icon">▼</span></th>
    <th>Estágio</th>
    <th>Modalidade</th>
    <th>Dias</th>
    <th class="day-num">1</th>
    <th class="day-num">2</th>
    <th class="day-num">3</th>
    <th class="day-num dom-num">4</th>
    <th class="day-num">5</th>
    <th class="day-num">6</th>
    <th class="day-num">7</th>
    <th class="day-num">8</th>
    <th class="day-num">9</th>
    <th class="day-num">10</th>
    <th class="day-num dom-num">11</th>
    <th class="day-num">12</th>
    <th class="day-num">13</th>
    <th class="day-num">14</th>
    <th class="day-num">15</th>
    <th class="day-num">16</th>
    <th class="day-num">17</th>
    <th class="day-num dom-num">18</th>
    <th class="day-num">19</th>
    <th class="day-num">20</th>
    <th class="day-num">21</th>
    <th class="day-num">22</th>
    <th class="day-num">23</th>
    <th class="day-num">24</th>
    <th class="day-num dom-num">25</th>
    <th class="day-num">26</th>
    <th class="day-num">27</th>
    <th class="day-num">28</th>
    <th class="day-num">29</th>
    <th class="day-num">30</th>
    <th class="day-num">31</th>
  </tr>

  <!-- ══════ GROUP: 8:00 ══════ -->
  <tr class="group-row">
    <td colspan="35">= 8:00</td>
  </tr>

  <!-- Alencássia Peres -->
  <tr class="data-row">
    <td class="nome">Alencássia Peres</td>
    <td class="estagio">W2 New</td>
    <td class="modalidade">Interactive</td>
    <td class="dias">2ª | 4ª</td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

  <!-- Aline Lucena -->
  <tr class="data-row">
    <td class="nome">Aline Lucena</td>
    <td class="estagio">ESP4</td>
    <td class="modalidade">Interactive</td>
    <td class="dias">2ª | 4ª</td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

  <!-- Ana Diniz -->
  <tr class="data-row">
    <td class="nome">Ana Diniz</td>
    <td class="estagio">W4 Old</td>
    <td class="modalidade">Interactive</td>
    <td class="dias">2ª | 4ª | 5ª</td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

  <!-- Ana Joyce -->
  <tr class="data-row">
    <td class="nome">Ana Joyce</td>
    <td class="estagio">W8</td>
    <td class="modalidade">Interactive</td>
    <td class="dias">2ª | 4ª</td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

  <!-- Ana Maria -->
  <tr class="data-row">
    <td class="nome">Ana Maria</td>
    <td class="estagio">W2 New</td>
    <td class="modalidade">Interactive</td>
    <td class="dias">2ª | 4ª</td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

  <!-- Ana Paula -->
  <tr class="data-row">
    <td class="nome">Ana Paula</td>
    <td class="estagio">W2 New</td>
    <td class="modalidade">Interactive</td>
    <td class="dias">2ª | 4ª | 6ª</td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

  <!-- ══════ GROUP: 15:00 ══════ -->
  <tr class="group-row">
    <td colspan="35">= 15:00</td>
  </tr>

  <!-- Ana Biatriz -->
  <tr class="data-row">
    <td class="nome">Ana Biatriz</td>
    <td class="estagio">NEXT GEN.</td>
    <td class="modalidade">Interactive</td>
    <td class="dias">2ª | 3ª | 4ª</td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

  <!-- ══════ GROUP: 17:00 ══════ -->
  <tr class="group-row">
    <td colspan="35">= 17:00</td>
  </tr>

  <!-- Antônio Moraes -->
  <tr class="data-row">
    <td class="nome">Antônio Moraes</td>
    <td class="estagio">W2 New</td>
    <td class="modalidade">Interactive</td>
    <td class="dias">2ª | 4ª</td>
    <td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell dom-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td><td class="day-cell"></td>
  </tr>

</table>
</div>
</body>
</html>
``` e há de verificar e estruturar os campos necessários no form do aluno.


