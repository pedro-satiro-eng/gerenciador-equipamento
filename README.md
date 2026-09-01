# gerenciador-equipamento codigo.gs
function doGet(e) {
  var html = HtmlService.createTemplateFromFile("Index");

  return html.evaluate()
    .setTitle("Gerenciador de Equipamentos")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}


// =====================================================
// CONFIGURAÇÃO INICIAL DOS PROFESSORES
// =====================================================
//
// EXECUTE ESTA FUNÇÃO UMA ÚNICA VEZ PARA CRIAR
// A ABA PROFESSORES.
//
// Depois disso, novos professores podem ser
// cadastrados DIRETAMENTE NA PLANILHA.
// =====================================================

function configurarProfessores() {

  var planilha = SpreadsheetApp.getActiveSpreadsheet();

  var aba = planilha.getSheetByName("PROFESSORES");

  if (!aba) {
    aba = planilha.insertSheet("PROFESSORES");
  }

  aba.clear();

  aba.getRange("A1:D1").setValues([[
    "PERÍODO",
    "EMAIL",
    "NOME",
    "ÁREA"
  ]]);

  var professores = [

    ["TARDE", "pamelapuccilo@professor.educacao.sp.gov.br", "Pamela Puccilo", "Biologia"],

    ["TARDE", "afaustobernardes@professor.educacao.sp.gov.br", "Fausto", "Desenvolvimento de Sistema"],

    ["TARDE", "katsantos@professor.educacao.sp.gov.br", "Katia", "Geografia"],

    ["TARDE", "elisaamorim@professor.educacao.sp.gov.br", "Maria Elisa", "Inglês"],

    ["TARDE", "paulomaia@professor.educacao.sp.gov.br", "Paulo", "Vendas"],

    ["TARDE", "ericalmeida@professor.educacao.sp.gov.br", "Eric", "Desenvolvimento de Sistema"],

    ["TARDE", "adrianaambrozio@professor.educacao.sp.gov.br", "Adriana", "Português"],

    ["TARDE", "marcella@professor.educacao.sp.gov.br", "Marcela", "Sociologia"],

    ["TARDE", "denisepedrita@professor.educacao.sp.gov.br", "Denise", "Português"],

    ["TARDE", "lauraferreira04@professor.educacao.sp.gov.br", "Laura", "Vendas"],

    ["TARDE", "paulaaparecidapaula@professor.educacao.sp.gov.br", "Paula", "Educação Física"],

    ["TARDE", "marciogoncalves1@professor.educacao.sp.gov.br", "Márcio", "Física"],

    ["TARDE", "merciasousa@professor.educacao.sp.gov.br", "Mercia", "Matemática"]

  ];

  aba.getRange(
    2,
    1,
    professores.length,
    4
  ).setValues(professores);

  aba.getRange("A1:D1")
    .setFontWeight("bold")
    .setBackground("#d9ead3");

  aba.setFrozenRows(1);

  aba.autoResizeColumns(1, 4);

  SpreadsheetApp.flush();

  return "✅ Aba PROFESSORES criada/configurada com sucesso!";
}


// =====================================================
// VALIDAR E-MAIL
// =====================================================

function validarEmailProfessor(email) {

  email =
    String(email || "")
      .trim()
      .toLowerCase();


  if (!email) {

    return {
      sucesso: false,
      mensagem:
        "Digite seu e-mail institucional."
    };
  }


  // O e-mail obrigatoriamente precisa conter
  // @professor

  if (
    email.indexOf("@professor") === -1
  ) {

    return {
      sucesso: false,
      mensagem:
        "⚠️ Utilize seu e-mail institucional @professor."
    };
  }


  var planilha =
    SpreadsheetApp
      .getActiveSpreadsheet();


  var aba =
    planilha.getSheetByName(
      "PROFESSORES"
    );


  if (!aba) {

    return {
      sucesso: false,
      mensagem:
        "❌ A aba PROFESSORES não foi configurada."
    };
  }


  var dados =
    aba.getDataRange()
      .getValues();


  for (
    var i = 1;
    i < dados.length;
    i++
  ) {

    var emailCadastrado =
      String(dados[i][1] || "")
        .trim()
        .toLowerCase();


    if (
      emailCadastrado === email
    ) {

      return {

        sucesso: true,

        periodo:
          String(dados[i][0] || "")
            .trim()
            .toUpperCase(),

        nome:
          String(dados[i][2] || "")
            .trim(),

        area:
          String(dados[i][3] || "")
            .trim()
      };
    }
  }


  return {

    sucesso: false,

    mensagem:
      "❌ Este e-mail não está autorizado a acessar o sistema."
  };
}


// =====================================================
// NOME DA ABA DO DIA
// =====================================================

function obterNomeAbaDoDia(periodo) {

  var data =
    new Date();


  var dia =
    String(
      data.getDate()
    ).padStart(2, "0");


  var mes =
    String(
      data.getMonth() + 1
    ).padStart(2, "0");


  var ano =
    String(
      data.getFullYear()
    ).slice(-2);


  return (
    dia +
    "." +
    mes +
    "." +
    ano +
    " - " +
    periodo
  );
}


// =====================================================
// OBTÉM / CRIA ABA DO DIA
// =====================================================

function obterAbaDoDia(
  planilha,
  periodo
) {

  var nomeAba =
    obterNomeAbaDoDia(
      periodo
    );


  var aba =
    planilha.getSheetByName(
      nomeAba
    );


  if (!aba) {

    aba =
      planilha.insertSheet(
        nomeAba
      );


    var cabecalhos = [

      "PROFESSOR",
      "OBSERVAÇÃO",
      "EQUIPAMENTOS",
      "NÚMEROS",
      "QUANTIDADE",
      "SITUAÇÃO",
      "HORÁRIO RETIRADA",
      "HORÁRIO DEVOLUÇÃO",
      "STATUS",
      "FALTANTES"

    ];


    aba
      .getRange(
        1,
        1,
        1,
        cabecalhos.length
      )
      .setValues([
        cabecalhos
      ]);


    aba
      .getRange("A1:J1")
      .setFontWeight("bold")
      .setBackground("#d9ead3");


    aba.setFrozenRows(1);
  }


  // ===================================================
  // CONVERSÃO DE ABA ANTIGA
  // ===================================================

  var cabecalhoB =
    String(
      aba.getRange("B1").getValue()
    )
      .trim()
      .toUpperCase();


  if (
    cabecalhoB === "EQUIPAMENTOS"
  ) {

    aba.insertColumnBefore(2);

    aba
      .getRange("B1")
      .setValue("OBSERVAÇÃO");

    aba
      .getRange("A1:J1")
      .setFontWeight("bold")
      .setBackground("#d9ead3");
  }


  return aba;
}


// =====================================================
// OBTÉM AS DUAS ABAS DO DIA
// =====================================================

function obterAbasDoDia(
  planilha
) {

  var abaManha =
    obterAbaDoDia(
      planilha,
      "MANHÃ"
    );


  var abaTarde =
    obterAbaDoDia(
      planilha,
      "TARDE"
    );


  return [
    abaManha,
    abaTarde
  ];
}


// =====================================================
// NORMALIZA QR CODE
// =====================================================

function normalizarCodigo(codigo) {

  codigo =
    String(codigo)
      .trim()
      .toUpperCase();


  // TABLET-01

  if (
    codigo.indexOf("TABLET-") === 0
  ) {

    var numero =
      parseInt(
        codigo.replace(
          "TABLET-",
          ""
        ),
        10
      );


    if (isNaN(numero)) {
      return null;
    }


    return {

      codigo:
        "T" +
        String(numero)
          .padStart(2, "0"),

      equipamento:
        "Tablet"
    };
  }


  // CHROMEBOOK-01

  if (
    codigo.indexOf("CHROMEBOOK-") === 0
  ) {

    var numero =
      parseInt(
        codigo.replace(
          "CHROMEBOOK-",
          ""
        ),
        10
      );


    if (isNaN(numero)) {
      return null;
    }


    return {

      codigo:
        "C" +
        String(numero)
          .padStart(2, "0"),

      equipamento:
        "Chromebook"
    };
  }


  // POSITIVO-01

  if (
    codigo.indexOf("POSITIVO-") === 0
  ) {

    var numero =
      parseInt(
        codigo.replace(
          "POSITIVO-",
          ""
        ),
        10
      );


    if (isNaN(numero)) {
      return null;
    }


    return {

      codigo:
        "P" +
        String(numero)
          .padStart(2, "0"),

      equipamento:
        "Positivo"
    };
  }


  // T01

  if (/^T\d+$/.test(codigo)) {

    var nT =
      parseInt(
        codigo.substring(1),
        10
      );


    return {

      codigo:
        "T" +
        String(nT)
          .padStart(2, "0"),

      equipamento:
        "Tablet"
    };
  }


  // C01

  if (/^C\d+$/.test(codigo)) {

    var nC =
      parseInt(
        codigo.substring(1),
        10
      );


    return {

      codigo:
        "C" +
        String(nC)
          .padStart(2, "0"),

      equipamento:
        "Chromebook"
    };
  }


  // P01

  if (/^P\d+$/.test(codigo)) {

    var nP =
      parseInt(
        codigo.substring(1),
        10
      );


    return {

      codigo:
        "P" +
        String(nP)
          .padStart(2, "0"),

      equipamento:
        "Positivo"
    };
  }


  return null;
}


// =====================================================
// ORDENAÇÃO
// =====================================================

function ordenarEquipamentos(
  lista
) {

  return lista.sort(
    function(a, b) {

      var ordem = {
        "T": 1,
        "C": 2,
        "P": 3
      };


      var tipoA =
        a.charAt(0);


      var tipoB =
        b.charAt(0);


      if (
        ordem[tipoA] !==
        ordem[tipoB]
      ) {

        return (
          ordem[tipoA] -
          ordem[tipoB]
        );
      }


      return (
        parseInt(
          a.substring(1),
          10
        ) -
        parseInt(
          b.substring(1),
          10
        )
      );
    }
  );
}


// =====================================================
// TRANSFORMA NÚMEROS EM ARRAY
// =====================================================

function transformarNumeros(
  texto
) {

  return String(
    texto || ""
  )
    .split(",")
    .map(
      function(item) {
        return item.trim();
      }
    )
    .filter(
      function(item) {
        return item !== "";
      }
    );
}


// =====================================================
// PROCURA DONO EM UMA ABA
// =====================================================

function encontrarDonoNaAba(
  aba,
  codigo
) {

  if (!aba) {
    return null;
  }


  var dados =
    aba.getDataRange()
      .getValues();


  for (
    var i = dados.length - 1;
    i >= 1;
    i--
  ) {

    var status =
      String(
        dados[i][8] || ""
      )
        .trim()
        .toUpperCase();


    if (
      status !== "PENDENTE" &&
      status !== "FALTANDO"
    ) {

      continue;
    }


    var numeros =
      transformarNumeros(
        dados[i][3]
      );


    if (
      numeros.includes(
        codigo
      )
    ) {

      return {

        professor:
          String(
            dados[i][0]
          ).trim(),

        linha:
          i + 1
      };
    }
  }


  return null;
}


// =====================================================
// PROCURA EQUIPAMENTO NAS DUAS ABAS
// =====================================================
//
// IMPORTANTE:
// Não importa se o professor é da manhã
// ou da tarde.
//
// O equipamento é considerado ocupado
// se estiver pendente em QUALQUER uma
// das duas abas.
// =====================================================

function encontrarDonoEquipamento(
  planilha,
  codigo
) {

  var abas =
    obterAbasDoDia(
      planilha
    );


  for (
    var i = 0;
    i < abas.length;
    i++
  ) {

    var dono =
      encontrarDonoNaAba(
        abas[i],
        codigo
      );


    if (dono) {

      return dono;
    }
  }


  return null;
}


// =====================================================
// PROCURA RETIRADA DO PROFESSOR
// NAS DUAS ABAS
// =====================================================

function encontrarRetiradaProfessor(
  planilha,
  nomeProfessor
) {

  var abas =
    obterAbasDoDia(
      planilha
    );


  for (
    var a = 0;
    a < abas.length;
    a++
  ) {

    var aba =
      abas[a];


    var dados =
      aba.getDataRange()
        .getValues();


    for (
      var i = dados.length - 1;
      i >= 1;
      i--
    ) {

      var professor =
        String(
          dados[i][0] || ""
        )
          .trim()
          .toUpperCase();


      var status =
        String(
          dados[i][8] || ""
        )
          .trim()
          .toUpperCase();


      if (
        professor ===
          nomeProfessor &&

        (
          status === "PENDENTE" ||
          status === "FALTANDO"
        )
      ) {

        return {

          aba: aba,

          linha:
            i + 1,

          dados:
            dados[i]
        };
      }
    }
  }


  return null;
}


// =====================================================
// CORES DOS NÚMEROS
// =====================================================

function aplicarCorNumeros(
  aba,
  linha,
  codigos
) {

  var texto =
    codigos.join(", ");


  var richText =
    SpreadsheetApp
      .newRichTextValue()
      .setText(texto);


  var posicao = 0;


  codigos.forEach(
    function(codigo) {

      var inicio =
        posicao;


      var fim =
        inicio +
        codigo.length;


      var cor =
        "#000000";


      if (
        codigo.charAt(0) === "T"
      ) {

        cor =
          "#8E44AD";

      }

      else if (
        codigo.charAt(0) === "C"
      ) {

        cor =
          "#1565C0";

      }

      else if (
        codigo.charAt(0) === "P"
      ) {

        cor =
          "#E67E22";
      }


      richText.setTextStyle(

        inicio,

        fim,

        SpreadsheetApp
          .newTextStyle()
          .setForegroundColor(cor)
          .setBold(true)
          .build()
      );


      posicao =
        fim + 2;
    }
  );


  aba
    .getRange(
      linha,
      4
    )
    .setRichTextValue(
      richText.build()
    );
}


// =====================================================
// CORES DO STATUS
// =====================================================

function aplicarCorStatus(
  aba,
  linha,
  status
) {

  var celula =
    aba.getRange(
      linha,
      9
    );


  if (
    status === "OK"
  ) {

    celula
      .setBackground("#93c47d")
      .setFontWeight("bold");

  }

  else if (
    status === "PENDENTE"
  ) {

    celula
      .setBackground("#FFD966")
      .setFontWeight("bold");

  }

  else if (
    status === "FALTANDO"
  ) {

    celula
      .setBackground("#E06666")
      .setFontWeight("bold");
  }
}


// =====================================================
// REGISTRAR LOTE
// =====================================================

function registrarLote(
  emailProfessor,
  acao,
  observacao,
  equipamentos
) {

  var lock =
    LockService
      .getScriptLock();


  try {

    lock.waitLock(10000);


    // =================================================
    // IDENTIFICA PROFESSOR PELO E-MAIL
    // =================================================

    var professor =
      validarEmailProfessor(
        emailProfessor
      );


    if (
      !professor.sucesso
    ) {

      return professor.mensagem;
    }


    var nomeProfessor =
      (
        professor.nome +
        " — " +
        professor.area
      )
        .trim()
        .toUpperCase();


    var periodo =
      professor.periodo;


    observacao =
      String(
        observacao || ""
      ).trim();


    if (
      !equipamentos ||
      equipamentos.length === 0
    ) {

      return (
        "⚠️ Nenhum equipamento foi lido."
      );
    }


    var planilha =
      SpreadsheetApp
        .getActiveSpreadsheet();


    var abaProfessor =
      obterAbaDoDia(
        planilha,
        periodo
      );


    // =================================================
    // NORMALIZA QR CODES
    // =================================================

    var codigos = [];

    var tipos = {};


    equipamentos.forEach(
      function(item) {

        var resultado =
          normalizarCodigo(
            item
          );


        if (!resultado) {
          return;
        }


        if (
          !codigos.includes(
            resultado.codigo
          )
        ) {

          codigos.push(
            resultado.codigo
          );


          tipos[
            resultado.equipamento
          ] = true;
        }
      }
    );


    if (
      codigos.length === 0
    ) {

      return (
        "⚠️ Nenhum QR Code válido foi encontrado."
      );
    }


    codigos =
      ordenarEquipamentos(
        codigos
      );


    var agora =
      new Date();


    var horario =
      Utilities.formatDate(
        agora,
        Session.getScriptTimeZone(),
        "HH'h'mm"
      );


    // =================================================
    // RETIRADA
    // =================================================

    if (
      acao === "Retirada"
    ) {


      // ===============================================
      // VERIFICA TODAS AS DUAS ABAS
      // ===============================================

      var equipamentosOcupados =
        [];


      codigos.forEach(
        function(codigo) {

          var dono =
            encontrarDonoEquipamento(
              planilha,
              codigo
            );


          if (dono) {

            equipamentosOcupados.push({

              codigo:
                codigo,

              professor:
                dono.professor
            });
          }
        }
      );


      if (
        equipamentosOcupados.length > 0
      ) {

        var mensagem =
          "⚠️ NÃO FOI POSSÍVEL REGISTRAR.\n\n" +
          "Os seguintes equipamentos já estão " +
          "com outro professor:\n\n";


        equipamentosOcupados.forEach(
          function(item) {

            mensagem +=
              "• " +
              item.codigo +
              " → " +
              item.professor +
              "\n";
          }
        );


        mensagem +=
          "\nDevolva o equipamento primeiro " +
          "antes de registrá-lo para outra pessoa.";


        return mensagem;
      }


      // ===============================================
      // PROFESSOR JÁ POSSUI RETIRADA?
      // ===============================================

      var retiradaExistente =
        encontrarRetiradaProfessor(
          planilha,
          nomeProfessor
        );


      if (
        retiradaExistente
      ) {

        return (
          "⚠️ " +
          nomeProfessor +
          " já possui uma retirada pendente."
        );
      }


      // ===============================================
      // REGISTRA NA ABA DO PERÍODO
      // ===============================================

      var equipamentosTexto =
        Object.keys(tipos)
          .join(", ");


      abaProfessor.appendRow([

        nomeProfessor,

        observacao,

        equipamentosTexto,

        codigos.join(", "),

        codigos.length,

        "Retirando",

        horario,

        "",

        "PENDENTE",

        ""

      ]);


      var linha =
        abaProfessor.getLastRow();


      aplicarCorStatus(
        abaProfessor,
        linha,
        "PENDENTE"
      );


      aplicarCorNumeros(
        abaProfessor,
        linha,
        codigos
      );


      return (
        "✅ Retirada registrada!\n\n" +
        "Professor: " +
        nomeProfessor +
        "\nEquipamentos: " +
        codigos.join(", ") +
        "\nQuantidade: " +
        codigos.length
      );
    }


    // =================================================
    // DEVOLUÇÃO
    // =================================================

    if (
      acao === "Devolução"
    ) {


      // ===============================================
      // PROCURA EM MANHÃ E TARDE
      // ===============================================

      var retirada =
        encontrarRetiradaProfessor(
          planilha,
          nomeProfessor
        );


      if (!retirada) {

        return (
          "⚠️ Nenhuma retirada pendente encontrada para " +
          nomeProfessor
        );
      }


      var abaRetirada =
        retirada.aba;


      var linhaEncontrada =
        retirada.linha;


      var dadosLinha =
        retirada.dados;


      // ===============================================
      // EQUIPAMENTOS RETIRADOS
      // ===============================================

      var numerosRetirados =
        transformarNumeros(
          dadosLinha[3]
        );


      // ===============================================
      // EQUIPAMENTOS ESTRANHOS
      // ===============================================

      var equipamentosEstranhos =
        codigos.filter(
          function(codigo) {

            return !numerosRetirados.includes(
              codigo
            );

          }
        );


      if (
        equipamentosEstranhos.length > 0
      ) {

        return (
          "⚠️ Este equipamento não pertence " +
          "à retirada de " +
          nomeProfessor +
          ":\n\n" +
          equipamentosEstranhos.join(", ") +
          "\n\nVerifique o QR Code."
        );
      }


      // ===============================================
      // DESCOBRE FALTANTES
      // ===============================================

      var faltantes =
        numerosRetirados.filter(
          function(codigo) {

            return !codigos.includes(
              codigo
            );

          }
        );


      var statusFinal =
        faltantes.length > 0
          ? "FALTANDO"
          : "OK";


      // ===============================================
      // HORÁRIO DEVOLUÇÃO
      // ===============================================

      abaRetirada
        .getRange(
          linhaEncontrada,
          8
        )
        .setValue(
          horario
        );


      // ===============================================
      // STATUS
      // ===============================================

      abaRetirada
        .getRange(
          linhaEncontrada,
          9
        )
        .setValue(
          statusFinal
        );


      // ===============================================
      // FALTANTES
      // ===============================================

      abaRetirada
        .getRange(
          linhaEncontrada,
          10
        )
        .setValue(

          faltantes.length > 0
            ? faltantes.join(", ")
            : ""

        );


      aplicarCorStatus(
        abaRetirada,
        linhaEncontrada,
        statusFinal
      );


      aplicarCorNumeros(
        abaRetirada,
        linhaEncontrada,
        numerosRetirados
      );


      // ===============================================
      // DEVOLUÇÃO COMPLETA
      // ===============================================

      if (
        statusFinal === "OK"
      ) {

        return (
          "✅ Devolução completa!\n\n" +
          "Todos os equipamentos foram devolvidos."
        );
      }


      // ===============================================
      // DEVOLUÇÃO INCOMPLETA
      // ===============================================

      return (
        "⚠️ Devolução registrada.\n\n" +
        "Equipamentos faltando:\n" +
        faltantes.join(", ")
      );
    }


    return "⚠️ Situação inválida.";


  } catch (erro) {

    return (
      "❌ Erro ao registrar:\n\n" +
      erro.message
    );

  } finally {

    lock.releaseLock();
  }
}
