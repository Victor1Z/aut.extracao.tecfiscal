"""Testes das funções que não dependem do navegador.

As funções de portal ficam fora de cobertura automatizada — a mitigação é o
comando `inspecionar` mais uma primeira rodada com `--limite 3`.
"""

import argparse
import csv

import pytest
from openpyxl import Workbook

import download_nfce as dn


def chave(sufixo: int) -> str:
    """Chave sintética de 44 dígitos, distinta por sufixo."""
    return f"{sufixo:044d}"


C1, C2, C3, C4 = chave(1), chave(2), chave(3), chave(4)


def planilha(tmp_path, valores, aba="Base", nome="base.xlsx"):
    """Gera um .xlsx mínimo com cabeçalho na linha 1 e `valores` a partir da 2."""
    livro = Workbook()
    planilha_ = livro.active
    planilha_.title = aba
    planilha_.append(["CHAVE_NFCE", "Situação Download"])
    for valor in valores:
        planilha_.append([valor])
    caminho = tmp_path / nome
    livro.save(caminho)
    return caminho


# ---------------------------------------------------------------------------
# CSV de controle
# ---------------------------------------------------------------------------


def test_controle_round_trip_preserva_todos_os_campos(tmp_path):
    caminho = tmp_path / "controle.csv"
    original = {
        C1: dn.Registro(C1, dn.STATUS_BAIXADO, 1, f"{C1}.pdf", "", "2026-08-10T09:00:00"),
        C2: dn.Registro(C2, dn.STATUS_ERRO, 2, "", "TimeoutError: sessão perdida", "2026-08-10T09:05:00"),
    }
    dn.gravar_controle(caminho, original)
    relido = dn.carregar_controle(caminho)

    assert relido == original


def test_controle_inexistente_carrega_vazio(tmp_path):
    assert dn.carregar_controle(tmp_path / "nao_existe.csv") == {}


def test_gravar_controle_escreve_cabecalho_esperado(tmp_path):
    caminho = tmp_path / "controle.csv"
    dn.gravar_controle(caminho, {C1: dn.Registro(C1)})

    with caminho.open(encoding=dn.ENCODING, newline="") as f:
        cabecalho = next(csv.reader(f))

    assert tuple(cabecalho) == dn.COLUNAS_CONTROLE


def test_gravar_controle_insiste_quando_o_destino_esta_bloqueado(tmp_path, monkeypatch):
    """OneDrive/antivírus bloqueiam o destino por instantes; insistir resolve."""
    caminho = tmp_path / "controle.csv"
    dn.gravar_controle(caminho, {C1: dn.Registro(C1, dn.STATUS_PENDENTE)})

    chamadas = {"n": 0}
    replace_real = dn.os.replace

    def replace_teimoso(origem, destino):
        chamadas["n"] += 1
        if chamadas["n"] < 3:
            raise PermissionError(5, "Acesso negado")
        return replace_real(origem, destino)

    monkeypatch.setattr(dn.os, "replace", replace_teimoso)
    monkeypatch.setattr(dn.time, "sleep", lambda _s: None)
    dn.gravar_controle(caminho, {C1: dn.Registro(C1, dn.STATUS_BAIXADO, 1, f"{C1}.pdf")})

    assert chamadas["n"] == 3
    assert dn.carregar_controle(caminho)[C1].status == dn.STATUS_BAIXADO


def test_gravar_controle_degrada_para_copia_se_o_replace_nunca_liberar(tmp_path, monkeypatch):
    caminho = tmp_path / "controle.csv"
    monkeypatch.setattr(dn.os, "replace", _sempre_negado)
    monkeypatch.setattr(dn.time, "sleep", lambda _s: None)
    dn.gravar_controle(caminho, {C1: dn.Registro(C1, dn.STATUS_BAIXADO, 1, f"{C1}.pdf")})

    assert dn.carregar_controle(caminho)[C1].status == dn.STATUS_BAIXADO
    assert not caminho.with_suffix(".csv.tmp").exists()


def _sempre_negado(_origem, _destino):
    raise PermissionError(5, "Acesso negado")


def test_sincronizacao_insere_chave_nova_como_pendente():
    controle = {}
    novas = dn.sincronizar_controle(controle, [C1, C2])

    assert novas == 2
    assert [r.status for r in controle.values()] == [dn.STATUS_PENDENTE] * 2
    assert controle[C1].atualizado_em != ""


def test_sincronizacao_preserva_chave_que_saiu_da_entrada():
    controle = {C1: dn.Registro(C1, dn.STATUS_BAIXADO, 1, f"{C1}.pdf")}
    dn.sincronizar_controle(controle, [C2])

    assert C1 in controle
    assert controle[C1].status == dn.STATUS_BAIXADO
    assert controle[C2].status == dn.STATUS_PENDENTE


def test_sincronizacao_nao_rebaixa_status_existente():
    controle = {C1: dn.Registro(C1, dn.STATUS_BAIXADO, 1, f"{C1}.pdf")}
    novas = dn.sincronizar_controle(controle, [C1])

    assert novas == 0
    assert controle[C1].status == dn.STATUS_BAIXADO


def test_fila_inclui_pendente_e_erro_e_exclui_baixado():
    controle = {
        C1: dn.Registro(C1, dn.STATUS_PENDENTE),
        C2: dn.Registro(C2, dn.STATUS_BAIXADO),
        C3: dn.Registro(C3, dn.STATUS_ERRO),
        C4: dn.Registro(C4, dn.STATUS_NAO_ENCONTRADO),
    }

    assert dn.montar_fila(controle, [C1, C2, C3, C4]) == [C1, C3]


def test_fila_com_retentar_inclui_nao_encontrado():
    controle = {
        C1: dn.Registro(C1, dn.STATUS_PENDENTE),
        C2: dn.Registro(C2, dn.STATUS_BAIXADO),
        C3: dn.Registro(C3, dn.STATUS_ERRO),
        C4: dn.Registro(C4, dn.STATUS_NAO_ENCONTRADO),
    }

    assert dn.montar_fila(controle, [C1, C2, C3, C4], retentar=True) == [C1, C3, C4]


def test_fila_ignora_chave_do_controle_ausente_da_entrada():
    controle = {C1: dn.Registro(C1, dn.STATUS_ERRO), C2: dn.Registro(C2, dn.STATUS_PENDENTE)}

    assert dn.montar_fila(controle, [C2]) == [C2]


def test_retomada_gera_fila_apenas_com_o_que_resta(tmp_path):
    caminho = tmp_path / "controle.csv"
    dn.gravar_controle(
        caminho,
        {
            C1: dn.Registro(C1, dn.STATUS_BAIXADO, 1, f"{C1}.pdf"),
            C2: dn.Registro(C2, dn.STATUS_PENDENTE),
        },
    )

    controle = dn.carregar_controle(caminho)
    chaves = [C1, C2, C3]
    dn.sincronizar_controle(controle, chaves)

    assert dn.montar_fila(controle, chaves) == [C2, C3]


def test_marcar_acumula_tentativas_e_trunca_o_erro():
    reg = dn.Registro(C1, dn.STATUS_ERRO, tentativas=2)
    reg.marcar(dn.STATUS_ERRO, erro="x" * 500, tentativas=2)

    assert reg.tentativas == 4
    assert len(reg.erro) == dn.LIMITE_MENSAGEM_ERRO


def test_contar_status_cobre_todos_os_status():
    controle = {C1: dn.Registro(C1, dn.STATUS_BAIXADO), C2: dn.Registro(C2, dn.STATUS_BAIXADO)}
    contagem = dn.contar_status(controle)

    assert contagem[dn.STATUS_BAIXADO] == 2
    assert contagem[dn.STATUS_PENDENTE] == 0


def test_ler_chaves_ignora_invalidas_e_duplicadas(tmp_path):
    caminho = tmp_path / "chaves.csv"
    dn.gravar_chaves(caminho, [C1, C2])
    with caminho.open("a", encoding=dn.ENCODING, newline="") as f:
        csv.writer(f).writerows([[C1], ["123"], [""]])

    assert dn.ler_chaves(caminho) == [C1, C2]


# ---------------------------------------------------------------------------
# Importação do .xlsm
# ---------------------------------------------------------------------------


def test_importacao_extrai_chaves_da_coluna_a_a_partir_da_linha_2(tmp_path):
    caminho = planilha(tmp_path, [C1, C2, C3])
    resultado = dn.importar_chaves_planilha(caminho)

    assert resultado.chaves == [C1, C2, C3]
    assert resultado.total_lido == 3
    assert resultado.invalidas == []
    assert resultado.duplicadas == 0


def test_importacao_descarta_chave_com_tamanho_errado_reportando_a_linha(tmp_path):
    caminho = planilha(tmp_path, [C1, "123", C2, "9" * 45])
    resultado = dn.importar_chaves_planilha(caminho)

    assert resultado.chaves == [C1, C2]
    assert [linha for linha, _ in resultado.invalidas] == [3, 5]


def test_importacao_remove_mascara_antes_de_validar(tmp_path):
    mascarada = f"{C1[:4]} {C1[4:8]}.{C1[8:]}"
    caminho = planilha(tmp_path, [mascarada])
    resultado = dn.importar_chaves_planilha(caminho)

    assert resultado.chaves == [C1]
    assert resultado.invalidas == []


def test_importacao_remove_duplicatas_preservando_a_ordem(tmp_path):
    caminho = planilha(tmp_path, [C2, C1, C2, C3, C1])
    resultado = dn.importar_chaves_planilha(caminho)

    assert resultado.chaves == [C2, C1, C3]
    assert resultado.duplicadas == 2


def test_importacao_ignora_linhas_vazias(tmp_path):
    caminho = planilha(tmp_path, [C1, None, "   ", C2])
    resultado = dn.importar_chaves_planilha(caminho)

    assert resultado.chaves == [C1, C2]
    assert resultado.total_lido == 2


def test_importacao_aba_inexistente_aborta_com_mensagem(tmp_path):
    caminho = planilha(tmp_path, [C1], aba="Outra")

    with pytest.raises(SystemExit, match="Base"):
        dn.importar_chaves_planilha(caminho, aba="Base")


def test_importacao_grava_chaves_csv_relegivel(tmp_path):
    caminho = planilha(tmp_path, [C1, "123", C2])
    resultado = dn.importar_chaves_planilha(caminho)
    saida = tmp_path / "chaves.csv"
    dn.gravar_chaves(saida, resultado.chaves)

    assert dn.ler_chaves(saida) == [C1, C2]


# ---------------------------------------------------------------------------
# Varredura de diretório
# ---------------------------------------------------------------------------


def test_varredura_encontra_arquivos_em_subpastas(tmp_path):
    (tmp_path / "2025" / "janeiro").mkdir(parents=True)
    (tmp_path / f"{C1}.pdf").write_text("x")
    (tmp_path / "2025" / "janeiro" / f"{C2}.pdf").write_text("x")

    mapa, _ = dn.varrer_destino(tmp_path)

    assert set(mapa) == {C1, C2}
    assert mapa[C2].parent.name == "janeiro"


def test_varredura_extrai_a_chave_de_nome_com_prefixo_e_sufixo(tmp_path):
    (tmp_path / f"NFCe_{C1}_v2.pdf").write_text("x")

    mapa, sem_chave = dn.varrer_destino(tmp_path)

    assert list(mapa) == [C1]
    assert sem_chave == []


def test_varredura_conta_como_orfao_arquivo_sem_chave_no_nome(tmp_path):
    (tmp_path / "relatorio.pdf").write_text("x")
    (tmp_path / "123.pdf").write_text("x")
    (tmp_path / f"{C1}.pdf").write_text("x")

    mapa, sem_chave = dn.varrer_destino(tmp_path)

    assert list(mapa) == [C1]
    assert sorted(p.name for p in sem_chave) == ["123.pdf", "relatorio.pdf"]


def test_varredura_casa_a_chave_independentemente_da_extensao(tmp_path):
    (tmp_path / f"{C1}.zip").write_text("x")
    (tmp_path / f"{C2}.xml").write_text("x")

    mapa, _ = dn.varrer_destino(tmp_path)

    assert mapa[C1].suffix == ".zip"
    assert mapa[C2].suffix == ".xml"


def test_varredura_de_destino_inexistente_devolve_vazio(tmp_path):
    mapa, sem_chave = dn.varrer_destino(tmp_path / "nao_existe")

    assert mapa == {} and sem_chave == []


# ---------------------------------------------------------------------------
# Reconciliação (comando conferir)
# ---------------------------------------------------------------------------


def test_conferir_marca_baixado_o_que_ja_esta_em_disco(tmp_path, capsys):
    pdfs = tmp_path / "pdfs" / "historico"
    pdfs.mkdir(parents=True)
    (pdfs / f"{C1}.pdf").write_text("x")
    (tmp_path / "pdfs" / f"{C4}.pdf").write_text("x")  # órfão: não está nas chaves

    caminho_chaves = tmp_path / "chaves.csv"
    caminho_controle = tmp_path / "controle.csv"
    dn.gravar_chaves(caminho_chaves, [C1, C2])

    args = argparse.Namespace(
        chaves=str(caminho_chaves), controle=str(caminho_controle), destino=str(tmp_path / "pdfs")
    )
    assert dn.comando_conferir(args) == 0

    controle = dn.carregar_controle(caminho_controle)
    assert controle[C1].status == dn.STATUS_BAIXADO
    assert controle[C1].arquivo == f"{C1}.pdf"
    assert controle[C2].status == dn.STATUS_PENDENTE
    assert "Arquivos órfãos             : 1" in capsys.readouterr().out


def test_extrair_chave_do_nome():
    assert dn.extrair_chave_do_nome(f"{C1}.pdf") == C1
    assert dn.extrair_chave_do_nome("sem chave.pdf") is None
    assert dn.extrair_chave_do_nome("2026-08-10.pdf") is None


def test_normalizar_chave_aceita_numero_grande_sem_notacao_cientifica():
    assert dn.normalizar_chave(float(10**20)) == "1" + "0" * 20
    assert dn.normalizar_chave(f" {C1} ") == C1


# ---------------------------------------------------------------------------
# Retry por chave
# ---------------------------------------------------------------------------


class PaginaFalsa:
    def __init__(self):
        self.gotos = []

    def goto(self, url, **kwargs):
        self.gotos.append(url)


def test_processar_chave_recarrega_a_pagina_a_cada_3_erros_seguidos(monkeypatch):
    """Erro de layout não pode travar o retry num loop que nunca se recupera."""
    monkeypatch.setattr(dn, "clicar", lambda *a, **k: (_ for _ in ()).throw(RuntimeError("boom")))
    monkeypatch.setattr(dn.time, "sleep", lambda _s: None)
    page = PaginaFalsa()

    resultado = dn.processar_chave(page, C1, None, 1000, tentativas=6)

    assert resultado.status == dn.STATUS_ERRO
    assert resultado.tentativas == 6
    # Erros seguidos: 1,2,3 -> recarrega; 4,5 sem recarregar (só volta a recarregar na 6ª,
    # mas essa é a última tentativa e o loop já termina sem reload).
    assert page.gotos == [dn.URL_NOTAS]


def test_processar_chave_nao_recarrega_se_tentativas_for_menor_que_3(monkeypatch):
    monkeypatch.setattr(dn, "clicar", lambda *a, **k: (_ for _ in ()).throw(RuntimeError("boom")))
    monkeypatch.setattr(dn.time, "sleep", lambda _s: None)
    page = PaginaFalsa()

    resultado = dn.processar_chave(page, C1, None, 1000, tentativas=2)

    assert resultado.status == dn.STATUS_ERRO
    assert page.gotos == []
    assert dn.normalizar_chave(None) == ""
