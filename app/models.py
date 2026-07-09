"""Data schema for the strategic review deck.

This mirrors the sections of the source presentation (cover, exec summary,
context charts, IPI-by-sector, line-of-business deep-dives, NPS, recovery
tracker, scenarios, recommended actions) without holding any real figures ---
all values here are just field definitions, populated at runtime from an
uploaded Excel workbook.
"""
from __future__ import annotations

from typing import List, Optional
from pydantic import BaseModel, Field


class Cover(BaseModel):
    title: str = ""
    subtitle: str = ""
    date_label: str = ""
    footer: str = ""


class TocEntry(BaseModel):
    number: str = ""
    section: str = ""
    description: str = ""
    page: str = ""


class IpiTier(BaseModel):
    label: str = ""
    bri_m: float = 0
    color_key: str = "on_track"  # on_track | cautious | at_risk


class LobStatus(BaseModel):
    lob: str = ""
    percent_on_track: float = 0
    bri_m: float = 0


class SavingsItem(BaseModel):
    category: str = ""
    amount_m: float = 0
    detail: str = ""


class ExecSummary(BaseModel):
    headline: str = ""
    subheadline: str = ""
    committed_bri_m: float = 0
    committed_label: str = ""
    savings_total_m: float = 0
    ipi_tiers: List[IpiTier] = Field(default_factory=list)
    benefit_status_by_lob: List[LobStatus] = Field(default_factory=list)
    savings_breakdown: List[SavingsItem] = Field(default_factory=list)


class GwpYear(BaseModel):
    year: str = ""
    bau_m: float = 0
    initiative_m: float = 0


class GwpByLob(BaseModel):
    lob: str = ""
    gwp_m: float = 0


class ContextGwp(BaseModel):
    headline: str = ""
    subheadline: str = ""
    years: List[GwpYear] = Field(default_factory=list)
    by_lob: List[GwpByLob] = Field(default_factory=list)
    by_lob_year_label: str = ""


class IpiSectorItem(BaseModel):
    lob: str = ""
    ipi: float = 0


class IpiSector(BaseModel):
    headline: str = ""
    subheadline: str = ""
    read_note: str = ""
    items: List[IpiSectorItem] = Field(default_factory=list)


class LobOverviewRow(BaseModel):
    lob: str = ""
    status_summary: str = ""
    bri_m: float = 0
    ipi: float = 0


class LobOverview(BaseModel):
    headline: str = ""
    narrative: str = ""
    rows: List[LobOverviewRow] = Field(default_factory=list)


class Initiative(BaseModel):
    lob: str = ""
    initiative: str = ""
    subtitle: str = ""
    ipi: float = 0
    ti: float = 0
    bri_committed_m: float = 0
    bri_actual_m: float = 0
    status: str = "on_track"  # on_track | cautious | at_risk
    comments: str = ""


class Project(BaseModel):
    lob: str = ""
    initiative: str = ""
    project: str = ""
    ipi: float = 0
    ti: float = 0
    comments: str = ""


class NpsLob(BaseModel):
    lob: str = ""
    actual: float = 0
    target: float = 0


class NpsSummary(BaseModel):
    headline: str = ""
    subheadline: str = ""
    company_actual: float = 0
    company_target: float = 0
    by_lob: List[NpsLob] = Field(default_factory=list)


class NpsDetailRow(BaseModel):
    lob: str = ""
    segment: str = ""
    actual: float = 0
    target: float = 0


class NpsDetail(BaseModel):
    headline: str = ""
    narrative: str = ""
    rows: List[NpsDetailRow] = Field(default_factory=list)


class RecoveryItem(BaseModel):
    rank: int = 0
    initiative: str = ""
    lob: str = ""
    item_type: str = ""  # growth | savings
    bri_m: float = 0
    status: str = "cautious"
    ipi: float = 0
    ti: float = 0
    root_cause: str = ""
    corrective_action: str = ""
    owner: str = ""


class RecoveryTracker(BaseModel):
    headline: str = ""
    subheadline: str = ""
    items: List[RecoveryItem] = Field(default_factory=list)


class Scenario(BaseModel):
    name: str = ""
    description: str = ""
    value_m: float = 0
    percent: float = 0


class Scenarios(BaseModel):
    headline: str = ""
    subheadline: str = ""
    committed_m: float = 0
    items: List[Scenario] = Field(default_factory=list)


class RecommendedAction(BaseModel):
    title: str = ""
    description: str = ""


class RecommendedActions(BaseModel):
    headline: str = ""
    subheadline: str = ""
    items: List[RecommendedAction] = Field(default_factory=list)


class Closing(BaseModel):
    message: str = "Thank you"
    footer: str = ""


class DeckData(BaseModel):
    cover: Cover = Field(default_factory=Cover)
    toc: List[TocEntry] = Field(default_factory=list)
    exec_summary: ExecSummary = Field(default_factory=ExecSummary)
    context_gwp: ContextGwp = Field(default_factory=ContextGwp)
    ipi_sector: IpiSector = Field(default_factory=IpiSector)
    lob_overview: LobOverview = Field(default_factory=LobOverview)
    initiatives: List[Initiative] = Field(default_factory=list)
    projects: List[Project] = Field(default_factory=list)
    nps_summary: NpsSummary = Field(default_factory=NpsSummary)
    nps_detail: NpsDetail = Field(default_factory=NpsDetail)
    recovery_tracker: RecoveryTracker = Field(default_factory=RecoveryTracker)
    scenarios: Scenarios = Field(default_factory=Scenarios)
    recommended_actions: RecommendedActions = Field(default_factory=RecommendedActions)
    closing: Closing = Field(default_factory=Closing)
