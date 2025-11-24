def test_imap_connection(host, port, username, password, use_ssl=True):
    """
    Tenta ligar ao servidor IMAP e retorna (ok, message, folders).
    folders é uma lista de nomes de pastas.
    """
    from imapclient import IMAPClient

    try:
        client = IMAPClient(host, port=port, ssl=use_ssl)
        client.login(username, password)
        folders = [f[2] for f in client.list_folders()]  # (flags, delimiter, name)
        client.logout()
        return True, "Ligação bem sucedida.", folders
    except Exception as e:
        return False, f"Erro na ligação: {e}", []
