import pandas as pd
import random

def create_dummy_data():
    # Define sample data
    data = {
        'Nº de Rastreio': [f'TRK{random.randint(10000, 99999)}' for _ in range(5)],
        'Número da NF-e': [f'{random.randint(100, 999)}' for _ in range(5)],
        'Nome do Destinatário': [
            'João Silva',
            'Maria Oliveira',
            'Carlos Souza',
            'Ana Pereira',
            'Pedro Santos'
        ],
        'SKU': [f'SKU-{random.randint(1, 50)}' for _ in range(5)],
        'Variação': ['P', 'M', 'G', 'XL', 'XXL'],
        'Qtd. do Produto': [1, 2, 1, 3, 1],
        'Link da Imagem': ['http://example.com/img.jpg'] * 5,
        'Extra Column 1': ['Ignore'] * 5,
        'Extra Column 2': ['Ignore'] * 5
    }

    # Create DataFrame
    df = pd.DataFrame(data)

    # Save to Excel
    output_file = 'pedidos_hoje.xlsx'
    df.to_excel(output_file, index=False)
    print(f"Created '{output_file}' with {len(df)} rows.")
    print("Example Tracking Numbers:")
    for tracking in df['Nº de Rastreio']:
        print(f" - {tracking}")

if __name__ == "__main__":
    create_dummy_data()
