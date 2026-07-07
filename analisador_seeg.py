# Criado por: marinaldosouza@gmail.com
import os
import pandas as pd
from pathlib import Path
import unicodedata
import re
import requests
from typing import Optional, List
from datetime import datetime

# --- Configurações Globais ---
# Desativa notação científica e exibe números com 2 casas decimais no terminal
pd.set_option('display.max_rows', 1000)
pd.options.display.float_format = '{:,.2f}'.format 

# Resolução dinâmica da Área de Trabalho real (Lendo do Registro do Windows para garantir 100% de precisão)
try:
    import winreg
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders")
    desktop_path, _ = winreg.QueryValueEx(key, "Desktop")
    winreg.CloseKey(key)
    DESKTOP_PATH = Path(os.path.expandvars(desktop_path))
except Exception:
    # Fallback de segurança 
    DESKTOP_PATH = Path.home() / 'Desktop'

WORK_DIR = DESKTOP_PATH / 'Python_df'
WORK_DIR.mkdir(parents=True, exist_ok=True)

FILE_URL = 'https://cdn3.gnarususercontent.com.br/2927-pandas-selecao-agrupamento-dados/1-SEEG10_GERAL-BR_UF_2022.10.27-FINAL-SITE.xlsx'
EXCEL_FILE_PATH = WORK_DIR / 'SEEG10_GERAL-BR.xlsx'

# --- Funções Auxiliares ---
def normalize_name(name: str) -> str:
    """Normaliza strings (remove acentos, espaços, caracteres especiais) para uso consistente."""
    normalized = unicodedata.normalize('NFKD', name).encode('ASCII', 'ignore').decode('utf-8')
    return re.sub(r'_+', '_', normalized.replace(' ', '_')).strip('_')

def get_user_numeric_input(prompt: str, allow_empty: bool = False, default: Optional[int] = None) -> Optional[int]:
    """Valida e retorna entrada numérica do usuário."""
    while True:
        user_input = input(prompt).strip()
        if allow_empty and not user_input:
            return default
        try:
            return int(user_input)
        except ValueError:
            print('Entrada inválida. Digite um número inteiro.')

def get_user_indices(prompt: str, max_index: int) -> List[int]:
    """Obtém uma lista de índices numéricos do usuário, validando o formato e o intervalo."""
    while True:
        indices_str = input(prompt).strip()
        if not indices_str:
            return []
        try:
            indices = [int(i.strip()) for i in indices_str.split(',') if i.strip()]
            if all(0 <= i < max_index for i in indices):
                return indices
            print(f'Um ou mais índices estão fora do intervalo (0 a {max_index - 1}). Tente novamente.')
        except ValueError:
            print('Erro: digite apenas números inteiros separados por vírgula.')

def find_sheet_dynamically(file_path: Path, target_name: str) -> str:
    """Busca o nome real da aba no Excel ignorando maiúsculas/minúsculas e espaços."""
    excel_file = pd.ExcelFile(file_path)
    available_sheets = excel_file.sheet_names
    
    target_clean = target_name.lower().replace(" ", "")
    for sheet in available_sheets:
        if target_clean in sheet.lower().replace(" ", ""):
            return sheet 
            
    print(f"\nAviso: Aba contendo '{target_name}' não encontrada. Abas disponíveis: {available_sheets}")
    print(f"Carregando a primeira aba padrão: '{available_sheets[0]}'\n")
    return available_sheets[0]

# --- Gestão de Arquivos, Download e Atualização ---
def _execute_download(url: str, dest_path: Path) -> None:
    """Função interna para realizar o download bruto do arquivo."""
    try:
        response = requests.get(url)
        response.raise_for_status()
        with open(dest_path, 'wb') as f:
            f.write(response.content)
        print('Download concluído com sucesso!')
    except requests.RequestException as e:
        print(f'Erro ao baixar o arquivo: {e}')
        exit(1)

def manage_database(url: str, dest_path: Path) -> None:
    """
    Verifica a existência do arquivo, gerencia versões e realiza backups
    com timestamp caso o usuário opte por atualizar a base.
    """
    if dest_path.exists():
        print(f"\n[!] A base de dados '{dest_path.name}' já existe na pasta.")
        while True:
            resposta = input("Deseja verificar e baixar uma versão atualizada da internet? (S/N): ").strip().lower()
            if resposta in ['s', 'sim']:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                backup_path = dest_path.with_name(f"{dest_path.stem}_{timestamp}{dest_path.suffix}")
                dest_path.rename(backup_path)
                print(f"\nArquivo antigo preservado como: {backup_path.name}")
                print("Iniciando download da nova base de dados...")
                _execute_download(url, dest_path)
                break
            elif resposta in ['n', 'nao', 'não']:
                print("Mantendo a base de dados atual.")
                break
            else:
                print("Opção inválida. Digite 'S' para sim ou 'N' para não.")
    else:
        print("\nBase de dados não encontrada localmente. Iniciando download automático...")
        _execute_download(url, dest_path)

def load_or_cache_dataframe(file_path: Path, target_sheet_name: str) -> pd.DataFrame:
    """
    Carrega dados do Excel e normaliza (removendo "Bunker") 
    mantendo os nomes originais das colunas intocados.
    Cria um cache em CSV para leitura rápida nas próximas execuções.
    """
    csv_path = WORK_DIR / f'{normalize_name(target_sheet_name)}_df.csv'
    
    # Se o CSV de cache existir e for mais recente que o arquivo Excel (ou seja, está atualizado)
    if csv_path.exists() and csv_path.stat().st_mtime > file_path.stat().st_mtime:
        print(f'Carregando dados normalizados do cache local: {csv_path.name}')
        return pd.read_csv(csv_path)
    
    print("Analisando e normalizando a base de dados bruta...")
    actual_sheet_name = find_sheet_dynamically(file_path, target_sheet_name)
    
    df = pd.read_excel(file_path, sheet_name=actual_sheet_name)
    
    # Normalização dos dados (apenas remoção de linhas específicas)
    if 'Bunker' in df.columns:
        df = df[df['Bunker'] != 'Sim']
    
    # Salva o resultado no CSV. Nomes de colunas permanecem 100% originais.
    df.to_csv(csv_path, index=False)
    print("Cache local atualizado com sucesso.")
    return df

# --- Funções de Visualização e Filtro ---
def display_dataframe_head(df: pd.DataFrame, n_lines: Optional[int] = None) -> None:
    """Exibe as primeiras N linhas de um DataFrame."""
    n_lines = n_lines if n_lines is not None else get_user_numeric_input('\nQuantas linhas deseja exibir? (pressione Enter para 100): ', allow_empty=True, default=100) or 100
    print(f'\nExibindo as primeiras {n_lines} linhas:')
    print(df.head(n_lines))

def view_or_select_columns(df: pd.DataFrame) -> None:
    """Permite visualizar o DataFrame completo ou selecionar colunas específicas."""
    print(f'\nTamanho do DataFrame: {df.shape[0]} linhas x {df.shape[1]} colunas\nColunas disponíveis:')
    for i, col in enumerate(df.columns):
        print(f'[{i}] {col} ({df[col].dtype})')
    
    indices = get_user_indices('\nDigite os índices das colunas desejadas separadas por vírgula (ou ENTER para ver todas): ', len(df.columns))
    
    if indices:
        selected_cols = [df.columns[i] for i in indices]
        print(f'Colunas selecionadas: {selected_cols}')
        display_dataframe_head(df[selected_cols])
    else:
        print('Nenhuma coluna selecionada. Exibindo o DataFrame completo.')
        display_dataframe_head(df)

def advanced_filter_data(df: pd.DataFrame) -> None:
    """Aplica filtros múltiplos em colunas selecionadas interativamente."""
    print("\n--- Aplicar Filtros (Básico/Avançado) ---")
    print("Escolha as colunas para iniciar o filtro:")
    for i, col in enumerate(df.columns):
        print(f"[{i}] {col}")

    base_indices = get_user_indices("\nDigite os índices das colunas para filtro (separadas por vírgula): ", len(df.columns))
    if not base_indices:
        print("Nenhuma coluna selecionada. Voltando ao menu.")
        return
    
    filtered_df = df.copy()
    for col_name in (df.columns[i] for i in base_indices):
        if col_name not in filtered_df.columns:
            continue
        
        unique_vals = sorted(filtered_df[col_name].dropna().unique())
        if not unique_vals:
            print(f"Não há valores únicos para a coluna '{col_name}'. Pulando.")
            continue
            
        print(f"\nValores disponíveis para '{col_name}':")
        for i, val in enumerate(unique_vals):
            print(f"[{i}] {val}")
        
        val_indices = get_user_indices(f"Escolha os índices dos valores para filtrar em '{col_name}' (vírgula para múltiplos, ENTER para pular): ", len(unique_vals))
        if val_indices:
            selected_vals = [unique_vals[i] for i in val_indices]
            filtered_df = filtered_df[filtered_df[col_name].isin(selected_vals)]
        
    print("\nResultado após filtros:")
    display_dataframe_head(filtered_df)

# --- Funções de Análise e Estatísticas ---
def calculate_statistics(df: pd.DataFrame) -> None:
    """Calcula estatísticas descritivas para colunas numéricas (anos) com base em filtros iterativos."""
    year_cols = [col for col in df.columns if re.fullmatch(r'\d{4}', str(col))]
    if not year_cols:
        print("Não foram encontradas colunas de anos para análise.")
        return

    print("\nColunas de anos disponíveis:")
    for i, year in enumerate(year_cols):
        print(f"[{i}] {year}")
        
    selected_year_indices = get_user_indices("Escolha os índices dos anos para análise (vírgula separada, ENTER para todos): ", len(year_cols))
    selected_year_cols = [year_cols[i] for i in selected_year_indices] if selected_year_indices else year_cols

    filtered_df = df.copy()

    # Filtro de Atividade
    print("\nColunas disponíveis para filtro de atividade:")
    for i, col in enumerate(df.columns):
        print(f"[{i}] {col}")
        
    act_idx = get_user_numeric_input("Escolha o índice da coluna para filtrar atividade (ou ENTER para pular): ", allow_empty=True)
    if act_idx is not None and 0 <= act_idx < len(df.columns):
        activity_col = df.columns[act_idx]
        act_vals = sorted(df[activity_col].dropna().unique())
        print(f"\nValores disponíveis para '{activity_col}':")
        for i, val in enumerate(act_vals):
            print(f"[{i}] {val}")
            
        selected_act_indices = get_user_indices("Escolha os índices das atividades (vírgula separada, ENTER para pular): ", len(act_vals))
        if selected_act_indices:
            selected_activities = [act_vals[i] for i in selected_act_indices]
            filtered_df = filtered_df[filtered_df[activity_col].isin(selected_activities)]

    # Filtro de Estado
    if 'Estado' in filtered_df.columns:
        state_vals = sorted(filtered_df['Estado'].dropna().unique())
        print("\nEstados disponíveis:")
        for i, state in enumerate(state_vals):
            print(f"[{i}] {state}")
            
        selected_state_indices = get_user_indices("Escolha os índices dos estados (vírgula separada, ENTER para pular): ", len(state_vals))
        if selected_state_indices:
            selected_states = [state_vals[i] for i in selected_state_indices]
            filtered_df = filtered_df[filtered_df['Estado'].isin(selected_states)]

    if filtered_df.empty:
        print("Nenhum dado encontrado com os filtros aplicados.")
        return

    print("\n--- Estatísticas para as Colunas de Anos Selecionados ---")
    try:
        stats_df = filtered_df[selected_year_cols].apply(pd.to_numeric, errors='coerce').dropna(how='all')
        if not stats_df.empty:
            stats_summary = stats_df.describe().map(lambda x: f"{x:,.2f}" if pd.notna(x) else "")
            print(stats_summary)
        else:
            print("Nenhum dado numérico disponível para calcular estatísticas com os filtros aplicados.")
    except Exception as e:
        print(f"Erro ao calcular estatísticas: {e}")

# --- Função Principal do Programa (Menu Interativo) ---
def main() -> None:
    """Ponto de entrada do programa."""
    print("="*60)
    print(" INICIALIZANDO O SISTEMA DE DADOS - GEE ESTADOS ")
    print("="*60)
    
    print(f"Caminho da Área de Trabalho configurado: {WORK_DIR.absolute()}")
    
    # 1. Gerencia o download e atualizações da base
    manage_database(FILE_URL, EXCEL_FILE_PATH)
    
    # 2. Carrega o DataFrame (cache CSV para performance e normalização)
    main_df = load_or_cache_dataframe(EXCEL_FILE_PATH, 'GEE Estados')

    menu_options = {
        '1': ("Visualizar/Selecionar Colunas", lambda: view_or_select_columns(main_df)),
        '2': ("Aplicar Filtros (Básico/Avançado)", lambda: advanced_filter_data(main_df)),
        '3': ("Estatísticas por Ano, Atividade e Estado", lambda: calculate_statistics(main_df)),
        '0': ("Sair", None)
    }

    while True:
        print('\n--- Menu Principal ---')
        for key, (desc, _) in menu_options.items():
            print(f'[{key}] {desc}')
        
        option = input('Digite sua opção: ').strip()
        if option == '0':
            print('Encerrando o programa. Até mais!')
            break
            
        action = menu_options.get(option)
        if action:
            action[1]()
        else:
            print('Opção inválida. Tente novamente.')

# --- Execução do Programa ---
if __name__ == '__main__':
    main()
