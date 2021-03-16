
WAGNER@WagnerBarros MINGW64 /c/GIT - GITHUB (master)
$ git clone https://github.com/Hamamellis/CursoExcel.git
Cloning into 'CursoExcel'... -- APÓS CRIARMOS NO GITHUB O REPOSITORIO (PASTA) copiamos o endereço dele e clonamos para a nossa máquina...
remote: Enumerating objects: 17, done.
remote: Counting objects: 100% (17/17), done.
remote: Compressing objects: 100% (16/16), done.
remote: Total 17 (delta 5), reused 0 (delta 0), pack-reused 0
Receiving objects: 100% (17/17), 4.52 MiB | 249.00 KiB/s, done.
Resolving deltas: 100% (5/5), done.

WAGNER@WagnerBarros MINGW64 /c/GIT - GITHUB (master)
$ cd CursoExcel

WAGNER@WagnerBarros MINGW64 /c/GIT - GITHUB/CursoExcel (main)
$ git status
On branch main
Your branch is up to date with 'origin/main'.

Changes not staged for commit:
  (use "git add <file>..." to update what will be committed)
  (use "git restore <file>..." to discard changes in working directory)
        modified:   "Script VBA Refor\303\247o Looping.vb"

Untracked files:
  (use "git add <file>..." to include in what will be committed)
        ~$APP - RASTREAMENTO 2020.xlsm

no changes added to commit (use "git add" and/or "git commit -a")

WAGNER@WagnerBarros MINGW64 /c/GIT - GITHUB/CursoExcel (main)
$ git commit -am "Uploading Archive"
[main 8f13876] Uploading Archive
 1 file changed, 1 insertion(+), 1 deletion(-)

WAGNER@WagnerBarros MINGW64 /c/GIT - GITHUB/CursoExcel (main)
$ git push
Enumerating objects: 5, done.
Counting objects: 100% (5/5), done.
Delta compression using up to 12 threads
Compressing objects: 100% (3/3), done.
Writing objects: 100% (3/3), 294 bytes | 294.00 KiB/s, done.
Total 3 (delta 2), reused 0 (delta 0), pack-reused 0
remote: Resolving deltas: 100% (2/2), completed with 2 local objects.
To https://github.com/Hamamellis/CursoExcel.git
   fb15bb9..8f13876  main -> main

WAGNER@WagnerBarros MINGW64 /c/GIT - GITHUB/CursoExcel (main)
$ git status
On branch main
Your branch is up to date with 'origin/main'.

nothing to commit, working tree clean

--SUBIMOS UM NOVO ARQUIVO DIRETO NO GITHUB, DIFERENTE NO REPOSITÓRIO LOCAL, ENTÃO FIZEMOS UM GIT PULL PARA ATUALIZAR NA MÁQUINA
WAGNER@WagnerBarros MINGW64 /c/GIT - GITHUB/CursoExcel (main)
$ git pull
remote: Enumerating objects: 12, done.
remote: Counting objects: 100% (12/12), done.
remote: Compressing objects: 100% (10/10), done.
remote: Total 10 (delta 2), reused 0 (delta 0), pack-reused 0
Unpacking objects: 100% (10/10), 106.23 KiB | 445.00 KiB/s, done.
From https://github.com/Hamamellis/CursoExcel
   8f13876..ffc4f15  main       -> origin/main
Updating 8f13876..ffc4f15
Fast-forward
 ... Tabuleiro Sequencial de N\303\272meros_1.xlsm" | Bin 0 -> 32913 bytes
 "Script VBA Refor\303\247o Looping.vb"             |   2 +-
 ...VBA da Consolida\303\247\303\243o de Contas.vb" | 144 +++++++++++++++++++++
 Sintaxe VBA para Limpar e Formatar.vb              |  93 +++++++++++++
 ...e VBA para a Execu\303\247\303\243o do Jogo.vb" |  97 ++++++++++++++
 ...e de Consolida\303\247\303\243o de Contas.xlsm" | Bin 0 -> 79003 bytes
 6 files changed, 335 insertions(+), 1 deletion(-)
 create mode 100644 "Aula Jogo de Tabuleiro Sequencial de N\303\272meros_1.xlsm"
 create mode 100644 "Sintaxe VBA da Consolida\303\247\303\243o de Contas.vb"
 create mode 100644 Sintaxe VBA para Limpar e Formatar.vb
 create mode 100644 "Sintaxe VBA para a Execu\303\247\303\243o do Jogo.vb"
 create mode 100644 "Template de Consolida\303\247\303\243o de Contas.xlsm"

WAGNER@WagnerBarros MINGW64 /c/GIT - GITHUB/CursoExcel (main)
$ git status
On branch main
Your branch is up to date with 'origin/main'.

nothing to commit, working tree clean

WAGNER@WagnerBarros MINGW64 /c/GIT - GITHUB/CursoExcel (main)
$
