from dotenv import load_dotenv
import os
import pandas as pd
from supabase import create_client

load_dotenv()


def create_supabase_client():
    supabase_url = os.getenv("SUPABASE_URL")
    supabase_key = os.getenv("SUPABASE_KEY")
    if not supabase_url or not supabase_key:
        raise EnvironmentError(
            "SUPABASE_URL และ SUPABASE_KEY ต้องตั้งค่าใน .env ก่อนรันโปรแกรม"
        )
    return create_client(supabase_url, supabase_key)


def fetch_buildings(client):
    response = client.table("surveys").select("building").execute()
    rows = getattr(response, "data", None) or []
    buildings = sorted(
        {row.get("building") for row in rows if row.get("building")}
    )
    return buildings


def fetch_survey_data(client, building):
    response = client.table("surveys").select("*").eq("building", building).execute()
    rows = getattr(response, "data", None) or []
    return pd.DataFrame(rows)


def fetch_traceroute_hops(client, survey_ids):
    if not survey_ids:
        return pd.DataFrame([])
    response = client.table("traceroute_hops").select("*").in_("survey_id", survey_ids).execute()
    rows = getattr(response, "data", None) or []
    return pd.DataFrame(rows)
