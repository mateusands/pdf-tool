"""Trabalho pesado fora da thread da UI.

Tkinter é single-threaded: operação longa no callback congela a janela, e tocar
em widget de dentro de outra thread corrompe o estado do Tk. Este módulo é o
único lugar do app que cria thread — ver a spec em `tests/test_background.py`.
"""

import threading


def executar_em_thread(root, trabalho, ao_terminar, ao_falhar) -> threading.Thread:
    """Roda `trabalho()` numa thread e volta para a UI por `root.after(0, …)`.

    `ao_terminar(resultado)` no sucesso, `ao_falhar(mensagem)` no erro.
    Devolve a thread (útil nos testes; a UI não precisa dela).
    """
    def tarefa():
        try:
            resultado = trabalho()
        except Exception as erro:
            # A mensagem é extraída AQUI, dentro do except: o Python apaga o
            # nome `erro` ao sair do bloco, e o callback só roda depois.
            mensagem = str(erro) or f"Falha inesperada ({type(erro).__name__})."
            root.after(0, lambda m=mensagem: ao_falhar(m))
            return
        root.after(0, lambda r=resultado: ao_terminar(r))

    thread = threading.Thread(target=tarefa, daemon=True)
    thread.start()
    return thread
