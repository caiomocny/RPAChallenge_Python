import RPAChallenge as RPApy
from tkinter import messagebox

# Passar message box para este arquivo (arquivo main)

if __name__ == '__main__':

    # Log Message - Start Process
    messagebox.showinfo("Star process", "O processo RPA Challenge será iniciado!")

    # Invocar o processo
    result = RPApy.main()

    # Log Message - End Process
    messagebox.showinfo("End process", "Processo finalizado com sucesso!\n"
                                       "Tempo de processamento:" + result)
