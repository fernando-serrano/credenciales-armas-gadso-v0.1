"""
Script para probar la normalización del Excel sin hacer validaciones de credenciales
"""
from dotenv import load_dotenv
import pandas as pd
import os
from datetime import datetime

load_dotenv()

EXCEL_DESNORMALIZADO = os.path.join("data", "credenciales-desnormalizado.xlsx")
EXCEL_NORMALIZADO = os.path.join("data", "credenciales-normalizado.xlsx")

def normalizar_excel_test():
    """
    Normaliza el Excel desnormalizado para pruebas
    """
    print("🔄 Leyendo Excel desnormalizado...")
    df = pd.read_excel(EXCEL_DESNORMALIZADO)
    
    # Hacer copia para no afectar original
    df_normalizado = df.copy()
    
    print(f"📋 Columnas encontradas: {df_normalizado.columns.tolist()}")
    print(f"📊 Total de registros: {len(df_normalizado)}\n")
    
    print("Vista previa de datos originales:")
    print(df_normalizado.head(2))
    print()
    
    # 1. Separar "marca temporal" en fecha y hora
    if 'marca temporal' in df_normalizado.columns:
        try:
            df_normalizado['marca temporal'] = pd.to_datetime(df_normalizado['marca temporal'], errors='coerce')
            df_normalizado['fecha'] = df_normalizado['marca temporal'].dt.date
            df_normalizado['hora'] = df_normalizado['marca temporal'].dt.strftime('%H:%M:%S')
            print("✅ Marca temporal separada en fecha y hora")
        except Exception as e:
            print(f"⚠️  Error al separar marca temporal: {e}")
    
    # 2. Normalizar DNI (agregar 0 adelante si tiene 7 dígitos)
    if 'dni' in df_normalizado.columns:
        df_normalizado['dni'] = df_normalizado['dni'].astype(str).str.strip()
        
        # Mostrar DNIs antes y después
        print("\n🔢 Normalización de DNI:")
        for idx, valor in enumerate(df_normalizado['dni']):
            if len(valor) == 7:
                dni_normalizado = '0' + valor
                print(f"  Registro {idx+1}: {valor} → {dni_normalizado}")
        
        df_normalizado['dni'] = df_normalizado['dni'].apply(
            lambda x: '0' + x if len(x) == 7 and x.isdigit() else x
        )
        print("✅ DNI normalizado (completado con 0 si necesario)")
    
    # 3. Convertir apellidos y nombres a mayúsculas
    nombre_cols = ['apelido paterno', 'apellido materno', 'nombres']
    
    print("\n📝 Normalización de nombres:")
    for col in nombre_cols:
        if col in df_normalizado.columns:
            print(f"  {col}: convertir a mayúsculas")
            df_normalizado[col] = df_normalizado[col].astype(str).str.strip().str.upper()
    
    # 4. Crear columna de nombre completo unificado
    if all(col in df_normalizado.columns for col in nombre_cols):
        df_normalizado['nombre_completo'] = (
            df_normalizado['apelido paterno'].fillna('') + ' ' +
            df_normalizado['apellido materno'].fillna('') + ' ' +
            df_normalizado['nombres'].fillna('')
        ).str.replace(r'\s+', ' ', regex=True).str.strip()
    
    print("✅ Nombres convertidos a mayúsculas")
    
    # 5. Agregar columna de estado (se rellenará después)
    if 'estado' not in df_normalizado.columns:
        df_normalizado['estado'] = ''
    
    print("\n✏️  Vista previa de datos normalizados:")
    print(df_normalizado[['dni', 'contraseña', 'apelido paterno', 'apellido materno', 'nombres', 'fecha', 'hora', 'estado']].head(2))
    
    # Guardar Excel normalizado
    df_normalizado.to_excel(EXCEL_NORMALIZADO, index=False)
    print(f"\n✅ Excel normalizado guardado en {EXCEL_NORMALIZADO}")
    
    return df_normalizado


if __name__ == "__main__":
    print("="*70)
    print("  PRUEBA DE NORMALIZACIÓN")
    print("="*70 + "\n")
    
    df_resultado = normalizar_excel_test()
    
    print("\n" + "="*70)
    print("  RESULTADO FINAL")
    print("="*70)
    print(f"Registros procesados: {len(df_resultado)}")
    print(f"Columnas del resultado: {df_resultado.columns.tolist()}")
    print("="*70)
