# Carpool Route Builder

This tool reads your Rides Sheet in Google Sheets, builds carpool groups, looks up home addresses from Full Address & Contact Info, computes traffic-aware pickup routes (outbound + return) to a final destination, writes blue, clickable Google Maps links back into the sheet (driver cell for Go, adjacent cell for Return), and saves a JSON summary.

### 1) Install & Use uv (virtualenv + deps)

We’ll use `uv` to create a local `.venv` and install dependencies from your `pyproject.toml`.

Basically, run the following commands in your terminal.

**macOS / Linux**

```bash
curl -LsSf https://astral.sh/uv/install.sh | sh

uv venv

source .venv/bin/activate

uv pip install -e .
```

**Windows** (note: this is untested)

```shell
irm https://astral.sh/uv/install.ps1 | iex

uv venv

. .\.venv\Scripts\Activate.ps1

uv pip install -e .
```

Basically, this does the following:
1. Installs UV
2. Creates a virtual environment called `.venv`
3. Activates the virtual environment (so when you use `python` it's using the python installed in `.venv`)
4. Installs the python libraries/dependencies for this project

#### Credentials

To run this project, you need Sheets access (service account) and Maps traffic (API key). 

To get this, I'll just share a file called `calicehockey-map-d5de75ad4b3d.json` with you. This gives you access to the service account I'm using which is how you read/edit the spreadsheet. I'll also give you a `.env` file that has a `GOOGLE_MAPS_API_KEY` which is what is used to get the live traffic data.

### 2) Running the script

To run the script, make sure:

1) the UV environment is activated (`source .venv/bin/activate`) 

2) You have both credential JSON files


Then, run the following:

```
python main.py
```

#### After running this, please double check whether it includes everyone by cross referencing the `results.json` file with the ride sheet. Failing to do so may result in people not getting picked up.

The ride sheet should now include two links for each driver: The first link which is a hyperlink on top of the driver's name should have directions to the rink. The second link should be to the right of the driver's name which has a link for the return route.

### Notes

I designed the code to make it very obvious if it can't find a driver's/passenger's address. If this happens, it should display a bright red error in the terminal displaying who's name it couldn't find the address of. To fix this, make sure their name matches the ride sheet. There may be other errors/bugs I'm not aware of, so after you run it, please do your due diligence and check `results.json` to ensure that what's displayed in that file matches the Google Sheet.