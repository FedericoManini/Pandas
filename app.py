import pandas as pd
from pathlib import Path
import resources.filename as rfile


def control(indirizzo):
    filexlsx = Path(indirizzo)
    df = pd.read_excel(filexlsx, skiprows=1)

    # Rinomina Colonne
    df.rename(columns=rfile.columnschange, inplace=True)

    # Settare Date e Ore
    df["DataInizio"] = pd.to_datetime(df["DataInizio"], format="%d/%m/%Y", errors='coerce')
    df["OrarioInizio"] = pd.to_datetime(df["OrarioInizio"], format="%H:%M:%S", errors='coerce').dt.time
    df["DataFine"] = pd.to_datetime(df["DataFine"], format="%d/%m/%Y", errors='coerce')
    df["OrarioFine"] = pd.to_datetime(df["OrarioFine"], format="%H:%M:%S", errors='coerce').dt.time
    df["Presidio"] = filexlsx.stem.replace("Cruscotto", "")
    # spostare tutti i campi "rotti" in un nuovo dataFrame che poi verrà gestito a parte
    errori = pd.DataFrame(df.loc[df["DataFine"].isna() | df['DataInizio'].isna() | df["OrarioFine"].isna() | df['OrarioInizio'].isna()])
    errori["Errore"] = "Orario o Data non inseriti o non validi"

    # cancellare dal df tutto ciò che ha NaN/NaT
    df = df.dropna(subset=["OrarioInizio", "OrarioFine", "DataInizio", "DataFine"])

    # filla i NAN con qualcosa se alcune colonne ne hanno bisogno
    df.fillna(rfile.fillna, inplace=True)

    # settaggio DateTimeInizio e DateTimeFine unendo i dati
    df["DateTimeInizio"] = pd.to_datetime(df["DataInizio"].astype(str)+' '+df['OrarioInizio'].astype(str))
    df["DateTimeFine"] = pd.to_datetime(df["DataFine"].astype(str)+' '+df['OrarioFine'].astype(str))

    # gestione errori
    diff_negative = pd.DataFrame(df[df["DateTimeFine"] <= df["DateTimeInizio"]])
    diff_negative["Errore"] = diff_negative.apply(lambda row: f"calcolati {int((row['DateTimeFine'] - row['DateTimeInizio']).total_seconds() / 60)} minuti di servizio", axis=1)

    sezioni_mancanti = pd.DataFrame(df[df["Cliente"].isna() | df["Direzione"].isna()])
    sezioni_mancanti["Errore"] = "Divisione o Cliente mancanti"

    errori = pd.concat([errori, diff_negative, sezioni_mancanti])

    errori["DataInizio"] = errori["DataInizio"].dt.date
    errori["DataFine"] = errori["DataFine"].dt.date
    errori.drop(columns=["DateTimeInizio", "DateTimeFine"], inplace=True)
    errori.index = errori.index + 3
    # errori.to_excel(f"{rfile.destinazione}\\{filexlsx.stem}_Verifica.xlsx")  # TODO uncomment when file needed
    return df

# percorso dal pc di lavoro


out = control(rfile.file2)

fuck = out.loc[(out["DateTimeInizio"] >= '2024-05-10') & (out["DateTimeInizio"] <= '2024-05-11')]

fuck.to_excel(f"{rfile.destinazione}\\Cruscottofiltrato.xlsx")
