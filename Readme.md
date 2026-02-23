# 📡 Monitor de Sensores en Tempo Real

Ola! Este proxecto é un programa en Python que **simula sensores** (temperatura, humidade, altura e presión) e garda os datos nun ficheiro Excel mentres os mostra nun panel web que se actualiza só. Nada complicado, prometido 🙂

---

## 🧐 Que fai este programa?

Imaxina que tes uns sensores nunha estación meteorolóxica. Cada 5 segundos envían datos como:

```
Temperatura: 21.3 °C  |  Humidade: 67.4 %  |  Altura: 845 m  |  Presión: 1013.2 hPa
```

O programa fai tres cousas á vez:
1. **Simula** eses sensores xerando datos aleatorios cada 5 segundos
2. **Garda** todos os datos nun ficheiro Excel (`sensores.xlsx`) con táboas e gráficas
3. **Mostra** un panel web que se refresca automaticamente para ver os datos en tempo real

---

## 🛠️ O que necesitas antes de empezar

- Un ordenador con **Windows, Mac ou Linux**
- Conexión a internet (só para a instalación)
- Ganas de aprender 🚀

---

## 1️⃣ Instalar Python

Python é a linguaxe de programación na que está escrito o programa. É gratuíto e moi popular.

### En Windows

1. Vai a [python.org/downloads](https://www.python.org/downloads/)
2. Fai clic no botón grande que di **"Download Python 3.x.x"**
3. Abre o ficheiro que se descargou
4. **MOI IMPORTANTE**: na primeira pantalla do instalador, marca a caixa que di **"Add Python to PATH"** — se non fas isto, non vai funcionar
5. Fai clic en **"Install Now"** e agarda a que remate

### En Mac

1. Vai a [python.org/downloads](https://www.python.org/downloads/)
2. Descarga o instalador para Mac
3. Ábreo e segue os pasos (seguinte, seguinte, instalar)

### En Linux

Probablemente xa o tes instalado. Podes comprobalo abrindo o terminal e escribindo:
```
python3 --version
```
Se che sae algo como `Python 3.x.x` xa está. Se non, o teu profe de informática pode axudarte.

---

## 2️⃣ Abrir o Terminal

O terminal é esa fiestra onde escribes comandos de texto. Non asustes, é máis sinxelo do que parece.

- **Windows**: preme `Windows + R`, escribe `cmd` e preme Enter
- **Mac**: preme `Cmd + Espazo`, escribe `Terminal` e preme Enter
- **Linux**: busca "Terminal" no menú de aplicacións

---

## 3️⃣ Descargar os ficheiros do proxecto

Descarga a carpeta do proxecto e gárdaa nun sitio que recuerdes (por exemplo, no Escritorio).

A carpeta ten que ter estes ficheiros dentro:
```
📁 proxecto_aquis/
   📄 stream_to_excel.py    ← o programa principal
   📄 visualizar.py         ← o panel web
   📄 Readme.md             ← isto que estás lendo
```

---

## 4️⃣ Navegar ata a carpeta no terminal

Temos que dicirlle ao terminal en que carpeta estamos. Usa o comando `cd`:

**En Windows** (se gardaches a carpeta no Escritorio):
```
cd C:\Users\TeuNome\Desktop\proxecto_aquis
```

**En Mac/Linux**:
```
cd ~/Desktop/proxecto_aquis
```

> 💡 **Truco**: en moitos ordenadores podes arrastrar a carpeta directamente á fiestra do terminal e escribe el camiño só.

Para comprobar que estás no sitio correcto, escribe `ls` (Mac/Linux) ou `dir` (Windows) e deberías ver os ficheiros `.py`.

---

## 5️⃣ Instalar as ferramentas necesarias

O programa usa unhas "librerías" — ferramentas extra para Python que hai que instalar unha soa vez. Escribe este comando no terminal:

```
pip3 install openpyxl streamlit pandas plotly
```

Verás moito texto mentres se descargan. Agarda a que remate (pode tardar un par de minutos). Se ao final non ves mensaxes de erro en vermello, perfecto ✅

> ⚠️ En Windows pode que o comando sexa `pip` en lugar de `pip3`. Se un non funciona, proba co outro.

---

## 6️⃣ Executar o programa

Necesitas ter **dúas ventás do terminal abertas** á vez, ambas na carpeta do proxecto.

### Terminal 1 — O que xera e garda os datos

```
python3 stream_to_excel.py
```

Deberías ver algo así cada 5 segundos:
```
============================================================
  Simulador de sensores  ->  Excel
  Ficheiro: sensores.xlsx   |   Intervalo: 5 s
  Ctrl+C para deter
============================================================
[SENSOR]  2024-03-15 10:23:01  T= 21.3°C  H= 67.4%  Alt= 845.0m  P= 1013.2hPa
[EXCEL]   #   1  T= 21.3°C  H= 67.4%  Alt= 845.0m  P= 1013.2hPa  -> gardado
```

Cada liña é unha nova lectura que se gardou no Excel.

### Terminal 2 — O panel web

Abre **outra ventá do terminal** (na mesma carpeta) e escribe:

```
streamlit run visualizar.py
```

O navegador debería abrirse só cunha páxina con todas as gráficas. Se non se abre, copia esta dirección no navegador:
```
http://localhost:8501
```

---

## 7️⃣ Que verás no panel web

**📊 Métricas na parte superior**
Catro tarxetiñas cos últimos valores de cada sensor. O número pequeno indica se subiu ou baixou respecto á lectura anterior.

**📈 Catro gráficas** (unha por parámetro)
- A **liña de cor** → o valor real de cada lectura
- A **liña gris discontinua** → a media de todas as lecturas ata agora
- A **zona sombreada** → o rango entre o mínimo e o máximo histórico

**📋 Táboa**
As últimas 20 lecturas en formato de táboa.

Todo actualízase automaticamente cada 5 segundos, sen tocar nada.

---

## 8️⃣ O ficheiro Excel

Na carpeta aparecerá un ficheiro `sensores.xlsx`. Podes abrilo con Excel ou LibreOffice Calc para ver os datos.

> ⚠️ **Ollo**: non o teñas aberto con Excel ao mesmo tempo que o programa está escribindo nel, pode dar erro. Péchao antes de que chegue a seguinte lectura.

Ten **6 follas**:

| Folla | Contido |
|---|---|
| **Rexistro** | Todas as lecturas xuntas nunha soa táboa |
| **Temperatura** | Seguimento da temperatura con mín, máx e media + gráfica |
| **Humidade** | O mesmo para a humidade |
| **Altura** | O mesmo para a altura |
| **Presión** | O mesmo para a presión |
| **Resumo** | Resumo xeral de todos os parámetros |

---

## 🛑 Como parar o programa

En calquera das dúas ventás do terminal, preme:

```
Ctrl + C
```

Isto detén o programa de xeito seguro e garda todo o que había.

---

## ❓ Problemas comúns

**"python3 non se recoñece como comando"**
→ En Windows proba con `python` en lugar de `python3`. Se tampouco funciona, reinstala Python marcando "Add Python to PATH".

**"No module named 'openpyxl'"** (ou calquera outro módulo)
→ As ferramentas non están instaladas. Repite o paso 5.

**O navegador non se abre só**
→ Copia `http://localhost:8501` na barra de enderezos do navegador.

**O Excel dá erro ao gardarse**
→ Téñelo aberto con Excel ao mesmo tempo. Péchao.

---

## 📁 Resumo dos ficheiros

```
proxecto_aquis/
│
├── stream_to_excel.py   → Xera os datos e actualiza o Excel
├── visualizar.py        → Panel web en tempo real
├── sensores.xlsx        → Créase automaticamente ao executar
└── Readme.md            → Esta guía
```

---

Calquera dúbida, pregunta sen medo. ¡Bo proveito! 🐍
