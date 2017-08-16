// assuming 50 is maximum number of answers, and a number of sheet structure details
// that could be parametrised.
//
// also assuming that answers do not contain ",", which is used as a separator
//
function SendEvaluationsToStudents(e) {  
  correcaoAutomaticaDasQuestoes()
  try {            
    var ss = SpreadsheetApp.getActiveSpreadsheet()
    var respostasdoform = ss.getSheetByName("Form Responses 1")
    var resumo = ss.getSheetByName("Resumo")
      
    var subject = "[ESS] resultados: " + ss.getName() 
       
    var questoes = respostasdoform.getSheetValues(1, 2, 1, 5) 
    var respostascorretas = resumo.getSheetValues(6, 14, 1, 4)
    var percentuaisdeacerto = resumo.getSheetValues(4, 14, 1, 4)
    var conceitosdaturma = resumo.getSheetValues(3, 10, 3, 3)
    var alunos = respostasdoform.getSheetValues(2, 2, 50,1)
    
    for (var i=0; i < alunos.length; i++) {
      if (alunos[i] != "") {        
        var respostasquestoes = respostasdoform.getSheetValues(i+2, 2, 1, 5) 
        var email = respostasquestoes[0][0]
        var acertos = resumo.getSheetValues(i+2, 7, 1, 1)
        var conceito = resumo.getSheetValues(i+2, 8, 1, 1)
        var message = introducao(email) + 
                      perguntasERespostas(questoes,respostasquestoes,respostascorretas,percentuaisdeacerto) + 
                      totalEConceitos(acertos,conceito,conceitosdaturma)  
        var x = message              
        MailApp.sendEmail(email, subject, message)
      }
    }
  } catch (e) {
    Logger.log(e.toString())
  }
}

function introducao(email) {
  var contact = ContactsApp.getContact(email)
  var name = "" 
  if (contact != null) name = contact.getFullName()
  var res = ""
  if (name == "") {
    res =  email
  } else {
    res = name.split(" ")[0] + " (" + email + ")"
  }
  res = res + ",\n\nObrigado por participar da avaliação. Ela serve principalmente para que "
            + "você avalie a sua dedicação e a forma de estudo do material em questão. Analise os resultados "
            + "e, se for o caso, pense em como melhorar o seu desempenho. Estou também à disposição para ajudar"
            + " no que for preciso.\n\nObrigado,\nPaulo\n\n"
  return res
}

function correcaoAutomaticaDasQuestoes() {    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var planilhaAlunos = spreadsheet.getSheetByName("Form Responses 1")
    var planilhaProfessor = spreadsheet.getSheetByName("Resumo")
    var numeroDeAlunos = planilhaAlunos.getLastRow() - 1
    var conjuntoDeRespostaAlunos = planilhaAlunos.getSheetValues(2, 3, numeroDeAlunos, 3)
    var conjuntoDeRespostasProfessor = planilhaProfessor.getSheetValues(6, 14, 1, 3)
    var gabarito = [conjuntoDeRespostasProfessor[0][0].toString().split(','),conjuntoDeRespostasProfessor[0][1].toString().split(','),conjuntoDeRespostasProfessor[0][2].toString().split(',')]
    
    for (var aluno = 0; aluno < numeroDeAlunos; aluno++) {
      computarEPreencherNotasDoAluno(aluno,conjuntoDeRespostaAlunos,gabarito,planilhaProfessor)   
    }
    preencherEspacosEmBranco(planilhaProfessor)
} 

function computarEPreencherNotasDoAluno(aluno, conjuntoDeRespostaAlunos, gabarito, planilhaProfessor) { 
    for (var questao = 0; questao < 3; questao++) {
      var respostaAlunoQuestaoAtual = conjuntoDeRespostaAlunos[aluno][questao].toString().split(',')
      var respostaProfessorQuestaoAtual = gabarito[questao]
      var nota = compararRespostasEComputarNotaDoAluno(respostaAlunoQuestaoAtual,respostaProfessorQuestaoAtual)
      preencherNotaDoAluno(planilhaProfessor,aluno,questao,nota)
    }
}

function preencherNotaDoAluno (planilhaProfessor, aluno, questao, nota) {
  planilhaProfessor.getRange(aluno + 2, questao + 2).setValue(nota)
}  

function compararRespostasEComputarNotaDoAluno(respostaAluno, respostaProfessor) {
    var erros = errosPorOmissao(respostaAluno, respostaProfessor) + errosPorInclusao(respostaAluno, respostaProfessor)
    return calcularNota(erros,respostaProfessor.length)
  }

// devaria ter assinalado mas não asinalou
function errosPorOmissao(respostaAluno, respostaProfessor) {
   return tamanhoDaDiferenca(respostaProfessor,respostaAluno) 
}

// assinalou mas não deveria ter asinalado
function errosPorInclusao(respostaAluno, respostaProfessor) {
   return tamanhoDaDiferenca(respostaAluno,respostaProfessor) 
}

function tamanhoDaDiferenca(a,b) {
   var diferenca = 0
   var tamanho = a.length
   for (var i = 0; i < tamanho; i++) {
      var e = a[i].trim()
      e = fixGoogleSheetsEncodingError(e)    
      if (!isIn(e,b)) {
          diferenca++
      }
   }
   return diferenca
}

  function isIn(alternativa,resposta) {
    var resultado = false
    var tamanho = resposta.length
    var j = 0
    while (j < tamanho && !resultado) {
       var alternativaResposta = resposta[j].trim()
       alternativaResposta = fixGoogleSheetsEncodingError(alternativaResposta)
       if (alternativaResposta == alternativa) {
         resultado = true
       }
       j = j + 1
    }
    return resultado
  }  

  function calcularNota(erradas,numeroAlternativasCorretasProfessor) {
    var nota
    if (erradas == 0) { 
      nota = 1 
    } else if (erradas == 1 && numeroAlternativasCorretasProfessor > 1) {
      nota = 0.5
    } else {
      nota = 0
    }  
    return nota
  }

  function fixGoogleSheetsEncodingError(txt) {
    if (txt.search('\xa0') != -1) {
      txt = txt.replace('\xa0',' ')
    }
    return txt
  }

  function preencherEspacosEmBranco(planilhaProfessor) {
    var respostas = planilhaProfessor.getSheetValues(1, 1, 50, 1)
    for(var i = 0; i < 50; i++) {
      if (respostas[i][0] === "") {
        planilhaProfessor.getRange(i + 1, 2).setValue("-")
        planilhaProfessor.getRange(i + 1, 3).setValue("-")
        planilhaProfessor.getRange(i + 1, 4).setValue("-")
        planilhaProfessor.getRange(i + 1, 5).setValue("-")
      }
    }
  }

function perguntasERespostas(questoes,respostasquestoes,respostascorretas,percentuaisdeacerto) {
  var res = ""
  for (var j=1; j<=3; j++) {
    res = res + "--------------------\nQuestão: " + questoes[0][j] + "\n\nSua resposta: " + 
          respostasquestoes[0][j] + "\n\nA resposta correta: " + respostascorretas[0][j-1] + 
          "\n\nPercentual de acerto da turma para essa questão: " + 
          percent(percentuaisdeacerto[0][j-1],1) + "\n\n"
  } 
  return res
}

function totalEConceitos(acertos,conceito,conceitosdaturma) {
  var res = "\n--------------------\nSeu total de acertos: " + acertos[0] + "\nSeu conceito para esta meta: " + 
            conceito[0] + "\n\nPercentuais da turma:\n" + percentuaisdaturma(conceitosdaturma) + 
            "\nMANA - Meta ainda não atingida \nMPA - Meta parcialmente atingida \nMA - Meta atingida\n"
  return res
}

function percentuaisdaturma(conceitosdaturma) {
  var res = ""
  for (var i=0; i<3; i++) {
      res = res + conceitosdaturma[i][0] + " - " + 
            numeroDeAlunos(conceitosdaturma[i][1]) + 
            percent(conceitosdaturma[i][2],1) + "\n"
  }   
  return res
}

function percent(number,precision) {
   return toFixed(number*100,precision) + "%"
}

function toFixed(value, precision) {
    var power = Math.pow(10, precision || 0)
    return String(Math.round(value * power) / power)
}

function numeroDeAlunos(n) {
  var res = n + " aluno"
  if (n >= 2) {
    res = res + "s"
  }
  res = res + " - "
  return res
}
