def login_user(username: str, password: str) -> Optional[Dict]:
    """Realiza login do usuário"""
    try:
        # Buscar usuário
        response = supabase.table("usuarios").select("*").eq("usuario", username).execute()
        
        if not response.data:
            return None
            
        user = response.data[0]
        
        # Verificar se usuário está ativo
        if not user.get('ativo', True):
            st.warning("⚠️ Usuário inativo. Entre em contato com o administrador.")
            return None
            
        # Verificar senha
        if verify_password(password, user.get('senha', '')):
            # Removido: atualização de ultimo_login porque a coluna não existe
            # supabase.table("usuarios").update({
            #     "ultimo_login": datetime.now().isoformat()
            # }).eq("id", user['id']).execute()
            
            # Registrar log (opcional)
            # registrar_log(user['id'], "LOGIN", f"Login realizado com sucesso")
            
            return user
        
        return None
        
    except Exception as e:
        st.error(f"Erro ao fazer login: {str(e)}")
        return None
