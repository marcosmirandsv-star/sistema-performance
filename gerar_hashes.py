# Criar o arquivo
cat > gerar_hashes.py << 'EOF'
import bcrypt

def hash_password(password):
    salt = bcrypt.gensalt(rounds=12)
    return bcrypt.hashpw(password.encode('utf-8'), salt).decode('utf-8')

print("=" * 60)
print("GERANDO HASHES PARA SUPABASE")
print("=" * 60)

hash_carine = hash_password('carine2026')
hash_marcos = hash_password('marcos2026')
hash_polyana = hash_password('polyana2026')

print(f"\nCARINE:")
print(f"Hash: {hash_carine}")

print(f"\nMARCOS:")
print(f"Hash: {hash_marcos}")

print(f"\nPOLYANA:")
print(f"Hash: {hash_polyana}")

print("\n" + "=" * 60)
print(f"""
UPDATE usuarios 
SET senha = '{hash_carine}' 
WHERE usuario = 'carine';

UPDATE usuarios 
SET senha = '{hash_marcos}' 
WHERE usuario = 'marcos';

UPDATE usuarios 
SET senha = '{hash_polyana}' 
WHERE usuario = 'polyana';
""")
EOF

# Executar o script
python gerar_hashes.py
