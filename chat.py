import openai  # pip install openai
import typer  # pip install "typer[all]"
from rich import print  # pip install rich
from rich.table import Table
import os






def main():

    openai.api_key = "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"
    # Contexto del asistente
    context = {"role": "system",
            "content": "Eres el asistente del Dr Peña y su equipo médico, tienes amplios conocimientos de medicina, fisiopatología, enfermedades, farmacología, fisiología humana y enfermeria. Tu funcion sera la de ofrecer guias y consejo a un equipo medico multidisciplinar"}
    messages = [context]
    prints()

    

    while True:
        try:
            content = __prompt()

            if content == "new":
                print("🆕 Nueva conversación creada")
                messages = [context]
                content = __prompt()
            elif content == 'save':
                with open('C:\\Users\\medico.RSD\\Desktop\\conversacion.txt', 'w', encoding='utf-8') as archivo:
                    for iteraccion in messages[1:]:
                        archivo.write(f"{iteraccion['content']}\n\n")
                    print('[bold blue]Archivo conversacion.txt guardado en el Escritorio[/bold blue]')            

            messages.append({"role": "user", "content": content})

            response = openai.ChatCompletion.create(
                model="gpt-4", messages=messages, max_tokens=2000, temperature=0.3)

            
            response_content = response.choices[0].message.content

            messages.append({"role": "assistant", "content": response_content})

            print(f"[bold green]> [/bold green] [green]{response_content}[/green]", '    ', '(({} tokens usados))'.format(response.usage.total_tokens))
        except openai.error.InvalidRequestError:
            print('[bold red]HAS ALCANZADO EL LIMITE MÁXIMO DE INTERACCIONES, LA CONVERSACION SERA RESETEADA, CONTEXTO PERDIDO[/bold red]')
            messages = [context]
            


def prints():
    print('''[bold red] ___                            _                                   _      _                     _         _
|  _|                          | |                                 (_)    | |                   (_)       | |
| |_   ___    ___   ___  _ __  | |_  _ __   ___    _ __   ___  ___  _   __| |  ___  _ __    ___  _   __ _ | |
|  _| / __|  / __| / _ \| '_ \ | __|| '__| / _ \  | '__| / _ \/ __|| | / _` | / _ \| '_ \  / __|| | / _` || |
| |  | (__  | (__ |  __/| | | || |_ | |   | (_) | | |   |  __/\__ \| || (_| ||  __/| | | || (__ | || (_| || |
|_|   \___|  \___| \___||_| |_| \__||_|    \___/  |_|    \___||___/|_| \__,_| \___||_| |_| \___||_| \__,_||_|


[/bold red]''')
    print("[bold green] **ASISTENTE SANITARIO PERSONAL DEL DR PEÑA Y SUS ENFERMERIT@S**[/bold green]")
    
    table = Table("Comando", "Descripción")
    table.add_row("exit", "Salir de la aplicación")
    table.add_row("new", "Crear una nueva conversación")
    table.add_row("clear", "Limpia la pantalla de conversaciones anteriores")
    table.add_row("save", "Guarda el contenido de la conversacion actual\nen el archivo 'conversacion.txt' localizado en el escritorio")
    table.add_row("Ctrl + C", "Copia el contenido seleccionado en el portapapeles")
    print(table)
    print("[bold red]\nIMPORTANTE!!:[/bold red] [blue]Escriba sus preguntas siempre en minusculas\n[/blue]")
    print('[blue]Utilice un lenguaje natural para formular su pregunta.[/blue]')
    print('[bold red]\nNUEVO GPT-4 IMPLEMENTADO, ES POSIBLE QUE SEA MAS LENTO EN CONTESTAR DE LO HABITUAL[/bold red]')

def __prompt() -> str:
    prompt = typer.prompt("\n>>> ")

    if prompt == "exit":
        exit = typer.confirm("¿Estás seguro?")
        if exit:
            print("¡Hasta luego!")
            raise typer.Abort()

        return __prompt()
    elif prompt == 'clear':
            os.system('cls' if os.name == 'nt' else 'clear')
            prints()

    return prompt


if __name__ == "__main__":
    
    os.system('cls' if os.name == 'nt' else 'clear')
    typer.run(main)
