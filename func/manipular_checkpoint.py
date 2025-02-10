import os

def carregar_checkpoint(arquivo):
    """Carrega os grupos ou subsets já processados do checkpoint."""
    if os.path.exists(arquivo):
        with open(arquivo, "r") as f:
            return set(f.read().splitlines())
    return set()
    
def salvar_checkpoint(arquivo, checkpoint_id):
    """Salva um grupo ou subset concluído no checkpoint."""
    print(f"✅ Salvando {checkpoint_id} no checkpoint...")
    with open(arquivo, "a") as f:
        f.write(f"{checkpoint_id}\n")  # Garante que o ID será salvo corretamente como string
