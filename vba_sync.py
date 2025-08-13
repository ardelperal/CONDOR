import os
import sys
import win32com.client
import tempfile
import shutil
import time
import subprocess

# --- Constants ---
acQuitSaveNone = 2 # Corresponds to acQuitSaveNone in VBA
# Access Application constants
acHidden = 0
acIcon = 1
acMaximized = 3

# --- Configuration ---
# Absolute path to the Access database
ACCDB_PATH = r"C:\Proyectos\CONDOR\back\Desarrollo\CONDOR.accdb"
# Absolute path to the source code directory
SRC_PATH = r"C:\Proyectos\CONDOR\src"
# Encoding for Access VBA files (typically Windows-1252 or cp1252 for Western languages)
ACCESS_ENCODING = 'cp1252'
# Encoding for Python/VS Code files (UTF-8 is standard)
VSCODE_ENCODING = 'utf-8'

def configure_silent_access(access_app):
    """Configure Access for completely silent operation."""
    try:
        access_app.Visible = False
        access_app.UserControl = False
        
        # Try to set DisplayAlerts (not available in all Access versions)
        try:
            access_app.DisplayAlerts = False
        except Exception:
            pass  # DisplayAlerts not available in this Access version
        
        # Additional silent configurations
        try:
            access_app.DoCmd.SetWarnings(False)
        except Exception:
            pass
            
        try:
            access_app.AutomationSecurity = 1  # msoAutomationSecurityLow
        except Exception:
            pass
            
        # Try to disable VBA project protection dialogs
        try:
            access_app.VBE.MainWindow.Visible = False
        except Exception:
            pass
            
    except Exception as e:
        print(f"Error en configuración silenciosa: {e}")

def get_access_app():
    """Returns an Access Application object configured for silent operation."""
    try:
        access_app = win32com.client.DispatchEx("Access.Application")
        configure_silent_access(access_app)
        return access_app
    except Exception as e:
        print(f"Error al iniciar Access: {e}")
        sys.exit(1)

def export_vba_modules():
    """Exports VBA modules from ACCDB to SRC_PATH."""
    print(f"Exportando módulos VBA de '{ACCDB_PATH}' a '{SRC_PATH}'...")
    access_app = get_access_app()

    try:
        access_app.OpenCurrentDatabase(ACCDB_PATH)
        vb_project = access_app.VBE.ActiveVBProject

        # Ensure SRC_PATH exists
        os.makedirs(SRC_PATH, exist_ok=True)

        for component in vb_project.VBComponents:
            file_name = component.Name
            file_extension = ""
            if component.Type == 1: # vbext_ct_StdModule
                file_extension = ".bas"
            elif component.Type == 2: # vbext_ct_ClassModule
                file_extension = ".cls"
            elif component.Type == 3: # vbext_ct_MSForm
                file_extension = ".frm"
            else:
                print(f"Saltando componente desconocido: {component.Name} (Tipo: {component.Type})")
                continue

            temp_export_path = os.path.join(tempfile.gettempdir(), file_name + file_extension)
            try:
                component.Export(temp_export_path)
                print(f"  Exportado temporalmente: {file_name}{file_extension}")

                # Read from temp file (Access's encoding) and write to SRC_PATH (VSCode's encoding)
                with open(temp_export_path, 'r', encoding=ACCESS_ENCODING, errors='replace') as f_in:
                    content = f_in.read()
                
                output_path = os.path.join(SRC_PATH, file_name + file_extension)
                with open(output_path, 'w', encoding=VSCODE_ENCODING) as f_out:
                    f_out.write(content)
                print(f"  Guardado en UTF-8: {output_path}")

            except Exception as e:
                print(f"Error al exportar/procesar '{file_name}{file_extension}': {e}")
            finally:
                if os.path.exists(temp_export_path):
                    os.remove(temp_export_path)

        print("Exportación completada.")

    except Exception as e:
        print(f"Error durante la exportación: {e}")
    finally:
        try:
            # Ensure silent shutdown
            access_app.UserControl = False
            
            # Try to disable alerts and warnings
            try:
                access_app.DisplayAlerts = False
            except:
                pass
            try:
                access_app.DoCmd.SetWarnings(False)
            except:
                pass
                
            # Save and quit without prompts
            try:
                access_app.DoCmd.Save()
            except:
                pass  # May fail if nothing to save
                
            access_app.Quit(acQuitSaveNone)
            # Give Access time to close properly
            time.sleep(1)
        except Exception as cleanup_e:
            print(f"Error durante la limpieza de Access: {cleanup_e}")
            # Force kill Access if it's still running
            try:
                subprocess.run(["taskkill", "/f", "/im", "msaccess.exe"], capture_output=True)
            except:
                pass

def import_vba_modules():
    """Imports VBA modules from SRC_PATH to ACCDB."""
    print(f"Importando módulos VBA de '{SRC_PATH}' a '{ACCDB_PATH}'...")
    access_app = get_access_app()
    
    # Additional silent configuration for import operations
    try:
        access_app.AutomationSecurity = 1  # msoAutomationSecurityLow
    except Exception as e:
        print(f"Advertencia: No se pudo establecer AutomationSecurity: {e}")

    try:
        access_app.OpenCurrentDatabase(ACCDB_PATH)
        vb_project = access_app.VBE.ActiveVBProject

        # --- Delete existing components first for a clean import ---
        print("Eliminando componentes VBA existentes en la base de datos...")
        components_to_delete = []
        for component in vb_project.VBComponents:
            print(f"  Preparando para eliminar: {component.Name}")
            components_to_delete.append(component.Name)
        
        for comp_name in components_to_delete:
            try:
                vb_project.VBComponents.Remove(vb_project.VBComponents(comp_name))
                print(f"  Eliminado correctamente: {comp_name}")
            except Exception as e:
                print(f"  ERROR al eliminar '{comp_name}': {e}")
        print("Componentes existentes eliminados.")
        # --- End delete ---

        for root, _, files in os.walk(SRC_PATH):
            for file_name in files:
                if file_name.endswith(('.bas', '.cls', '.frm')):
                    src_file_path = os.path.join(root, file_name)
                    print(f"  Procesando: {src_file_path}")

                    try:
                        # Read from SRC_PATH (VSCode's encoding)
                        with open(src_file_path, 'r', encoding=VSCODE_ENCODING) as f_in:
                            content = f_in.read()
                        
                        # Write to temp file (Access's encoding) for import
                        temp_import_path = os.path.join(tempfile.gettempdir(), file_name)
                        with open(temp_import_path, 'w', encoding=ACCESS_ENCODING) as f_out:
                            f_out.write(content)
                        print(f"  Guardado temporalmente en ANSI: {temp_import_path}")

                        # Import with error handling for dialogs
                        try:
                            # Force silent import by temporarily disabling user interaction
                            original_user_control = access_app.UserControl
                            access_app.UserControl = False
                            
                            # Import the component
                            vb_project.VBComponents.Import(temp_import_path)
                            print(f"  Importado correctamente: {file_name}")
                            
                            # Restore original user control
                            access_app.UserControl = original_user_control
                            
                        except Exception as import_e:
                            print(f"  ERROR al importar '{file_name}': {import_e}")
                            # Try alternative import method
                            try:
                                print(f"  Intentando método alternativo para: {file_name}")
                                # Use DoCmd.RunCommand to import without dialogs
                                access_app.DoCmd.TransferText(0, "", file_name.split('.')[0], temp_import_path, False)
                                print(f"  Importado con método alternativo: {file_name}")
                            except Exception as alt_e:
                                print(f"  ERROR en método alternativo para '{file_name}': {alt_e}")

                    except Exception as e:
                        print(f"  ERROR general al procesar '{file_name}': {e}")
                    finally:
                        if 'temp_import_path' in locals() and os.path.exists(temp_import_path):
                            os.remove(temp_import_path)
        
        print("Importación completada.")

    except Exception as e:
        print(f"Error durante la importación: {e}")
    finally:
        try:
            # Ensure silent shutdown
            access_app.UserControl = False
            
            # Try to disable alerts and warnings
            try:
                access_app.DisplayAlerts = False
            except:
                pass
            try:
                access_app.DoCmd.SetWarnings(False)
            except:
                pass
                
            # Save and quit without prompts
            try:
                access_app.DoCmd.Save()
            except:
                pass  # May fail if nothing to save
                
            access_app.Quit(acQuitSaveNone)
            # Give Access time to close properly
            time.sleep(1)
        except Exception as cleanup_e:
            print(f"Error durante la limpieza de Access: {cleanup_e}")
            # Force kill Access if it's still running
            try:
                subprocess.run(["taskkill", "/f", "/im", "msaccess.exe"], capture_output=True)
            except:
                pass

def main():
    if len(sys.argv) < 2:
        print("Uso: python vba_sync.py [export|import]")
        sys.exit(1)

    command = sys.argv[1].lower()

    if command == "export":
        export_vba_modules()
    elif command == "import":
        import_vba_modules()
    else:
        print("Comando no reconocido. Uso: python vba_sync.py [export|import]")
        sys.exit(1)

if __name__ == "__main__":
    main()
