"""
Servizio per la generazione di report delle transazioni rifiutate.
Gestisce la creazione di report mensili, analisi e statistiche.
"""

import logging
from datetime import datetime
from decimal import Decimal
from typing import List, Dict, Any, Optional
from pathlib import Path
import uuid

from models.transaction_reporting import (
    Transaction, 
    RejectionReport, 
    MonthlyReportConfig, 
    ProcessingResult
)


class ReportGenerationService:
    """Servizio per la generazione di report delle transazioni rifiutate."""
    
    def __init__(self):
        """Inizializza il servizio."""
        self.logger = logging.getLogger(__name__)
    
    def generate_monthly_report(
        self, 
        transactions: List[Transaction], 
        config: MonthlyReportConfig
    ) -> RejectionReport:
        """
        Genera un report mensile delle transazioni rifiutate.
        
        Args:
            transactions: Lista delle transazioni da analizzare
            config: Configurazione per il report
            
        Returns:
            Report delle transazioni rifiutate
        """
        try:
            self.logger.info(f"Generazione report mensile per {config.year}/{config.month:02d}")
            
            # Filtra transazioni per periodo
            period_start = config.get_period_start()
            period_end = config.get_period_end()
            
            period_transactions = [
                t for t in transactions
                if period_start <= t.transaction_date < period_end
            ]
            
            self.logger.info(f"Trovate {len(period_transactions)} transazioni nel periodo")
            
            # Filtra transazioni rifiutate
            rejected_transactions = self._filter_rejected_transactions(
                period_transactions, config
            )
            
            self.logger.info(f"Trovate {len(rejected_transactions)} transazioni rifiutate")
            
            # Calcola statistiche
            total_amount = sum(t.amount for t in period_transactions)
            rejected_amount = sum(t.amount for t in rejected_transactions)
            
            # Analisi per tipo di transazione
            rejection_by_type = self._analyze_by_transaction_type(rejected_transactions)
            
            # Analisi per motivo di rifiuto
            rejection_by_reason = self._analyze_by_rejection_reason(rejected_transactions)
            
            # Crea il report
            report = RejectionReport(
                report_id=str(uuid.uuid4()),
                generation_date=datetime.now(),
                period_start=period_start,
                period_end=period_end,
                total_transactions=len(period_transactions),
                rejected_transactions=len(rejected_transactions),
                rejection_rate=0.0,  # Calcolato automaticamente in __post_init__
                total_amount=total_amount,
                rejected_amount=rejected_amount,
                rejection_by_type=rejection_by_type,
                rejection_by_reason=rejection_by_reason,
                transactions=rejected_transactions
            )
            
            self.logger.info(f"Report generato: {report.rejection_rate:.2f}% rejection rate")
            return report
            
        except Exception as e:
            self.logger.error(f"Errore nella generazione del report: {e}")
            raise
    
    def generate_summary_statistics(self, report: RejectionReport) -> Dict[str, Any]:
        """
        Genera statistiche riassuntive per il report.
        
        Args:
            report: Report delle transazioni rifiutate
            
        Returns:
            Dizionario con le statistiche riassuntive
        """
        try:
            # Statistiche di base
            stats = {
                "period": {
                    "start": report.period_start.strftime("%Y-%m-%d"),
                    "end": report.period_end.strftime("%Y-%m-%d"),
                    "days": (report.period_end - report.period_start).days
                },
                "transactions": {
                    "total": report.total_transactions,
                    "rejected": report.rejected_transactions,
                    "approved": report.total_transactions - report.rejected_transactions,
                    "rejection_rate": report.rejection_rate
                },
                "amounts": {
                    "total": float(report.total_amount),
                    "rejected": float(report.rejected_amount),
                    "approved": float(report.total_amount - report.rejected_amount),
                    "average_transaction": float(report.total_amount / report.total_transactions) if report.total_transactions > 0 else 0,
                    "average_rejected": float(report.rejected_amount / report.rejected_transactions) if report.rejected_transactions > 0 else 0
                },
                "rejection_analysis": {
                    "by_type": report.rejection_by_type,
                    "by_reason": report.rejection_by_reason,
                    "top_reasons": self._get_top_rejection_reasons(report.rejection_by_reason, 5)
                }
            }
            
            # Analisi temporale
            stats["temporal_analysis"] = self._analyze_temporal_patterns(report.transactions)
            
            # Analisi per importo
            stats["amount_analysis"] = self._analyze_amount_patterns(report.transactions)
            
            return stats
            
        except Exception as e:
            self.logger.error(f"Errore nella generazione delle statistiche: {e}")
            raise
    
    def compare_with_previous_period(
        self, 
        current_report: RejectionReport, 
        previous_report: RejectionReport
    ) -> Dict[str, Any]:
        """
        Confronta il report corrente con quello del periodo precedente.
        
        Args:
            current_report: Report del periodo corrente
            previous_report: Report del periodo precedente
            
        Returns:
            Dizionario con il confronto tra i periodi
        """
        try:
            comparison = {
                "current_period": {
                    "start": current_report.period_start.strftime("%Y-%m-%d"),
                    "end": current_report.period_end.strftime("%Y-%m-%d"),
                    "rejection_rate": current_report.rejection_rate,
                    "total_transactions": current_report.total_transactions,
                    "rejected_transactions": current_report.rejected_transactions
                },
                "previous_period": {
                    "start": previous_report.period_start.strftime("%Y-%m-%d"),
                    "end": previous_report.period_end.strftime("%Y-%m-%d"),
                    "rejection_rate": previous_report.rejection_rate,
                    "total_transactions": previous_report.total_transactions,
                    "rejected_transactions": previous_report.rejected_transactions
                },
                "changes": {
                    "rejection_rate_change": current_report.rejection_rate - previous_report.rejection_rate,
                    "transaction_volume_change": current_report.total_transactions - previous_report.total_transactions,
                    "rejected_volume_change": current_report.rejected_transactions - previous_report.rejected_transactions,
                    "amount_change": float(current_report.total_amount - previous_report.total_amount)
                }
            }
            
            # Calcola percentuali di variazione
            if previous_report.total_transactions > 0:
                comparison["changes"]["transaction_volume_pct"] = (
                    (current_report.total_transactions - previous_report.total_transactions) / 
                    previous_report.total_transactions * 100
                )
            
            if previous_report.rejected_transactions > 0:
                comparison["changes"]["rejected_volume_pct"] = (
                    (current_report.rejected_transactions - previous_report.rejected_transactions) / 
                    previous_report.rejected_transactions * 100
                )
            
            return comparison
            
        except Exception as e:
            self.logger.error(f"Errore nel confronto tra periodi: {e}")
            raise
    
    def _filter_rejected_transactions(
        self, 
        transactions: List[Transaction], 
        config: MonthlyReportConfig
    ) -> List[Transaction]:
        """Filtra le transazioni rifiutate secondo la configurazione."""
        rejected_statuses = ["REJECTED", "FAILED", "DENIED", "DECLINED"]
        
        if config.include_pending:
            rejected_statuses.append("PENDING")
        
        if config.include_failed:
            rejected_statuses.extend(["ERROR", "TIMEOUT"])
        
        return [
            t for t in transactions
            if t.status.upper() in rejected_statuses
        ]
    
    def _analyze_by_transaction_type(self, transactions: List[Transaction]) -> Dict[str, int]:
        """Analizza le transazioni rifiutate per tipo."""
        type_counts = {}
        for t in transactions:
            type_counts[t.transaction_type] = type_counts.get(t.transaction_type, 0) + 1
        return type_counts
    
    def _analyze_by_rejection_reason(self, transactions: List[Transaction]) -> Dict[str, int]:
        """Analizza le transazioni rifiutate per motivo."""
        reason_counts = {}
        for t in transactions:
            reason = t.rejection_reason or "UNKNOWN"
            reason_counts[reason] = reason_counts.get(reason, 0) + 1
        return reason_counts
    
    def _get_top_rejection_reasons(self, reasons: Dict[str, int], limit: int) -> List[Dict[str, Any]]:
        """Ottiene i primi N motivi di rifiuto per frequenza."""
        sorted_reasons = sorted(reasons.items(), key=lambda x: x[1], reverse=True)
        return [
            {"reason": reason, "count": count}
            for reason, count in sorted_reasons[:limit]
        ]
    
    def _analyze_temporal_patterns(self, transactions: List[Transaction]) -> Dict[str, Any]:
        """Analizza i pattern temporali delle transazioni rifiutate."""
        if not transactions:
            return {}
        
        # Analisi per ora del giorno
        hourly_counts = {}
        daily_counts = {}
        
        for t in transactions:
            hour = t.transaction_date.hour
            day = t.transaction_date.strftime("%A")
            
            hourly_counts[hour] = hourly_counts.get(hour, 0) + 1
            daily_counts[day] = daily_counts.get(day, 0) + 1
        
        return {
            "hourly_distribution": hourly_counts,
            "daily_distribution": daily_counts,
            "peak_hour": max(hourly_counts, key=hourly_counts.get) if hourly_counts else None,
            "peak_day": max(daily_counts, key=daily_counts.get) if daily_counts else None
        }
    
    def _analyze_amount_patterns(self, transactions: List[Transaction]) -> Dict[str, Any]:
        """Analizza i pattern degli importi delle transazioni rifiutate."""
        if not transactions:
            return {}
        
        amounts = [float(t.amount) for t in transactions]
        
        # Categorizzazione per range di importo
        ranges = {
            "0-10": 0,
            "10-50": 0,
            "50-100": 0,
            "100-500": 0,
            "500-1000": 0,
            "1000+": 0
        }
        
        for amount in amounts:
            if amount <= 10:
                ranges["0-10"] += 1
            elif amount <= 50:
                ranges["10-50"] += 1
            elif amount <= 100:
                ranges["50-100"] += 1
            elif amount <= 500:
                ranges["100-500"] += 1
            elif amount <= 1000:
                ranges["500-1000"] += 1
            else:
                ranges["1000+"] += 1
        
        return {
            "total_amount": sum(amounts),
            "average_amount": sum(amounts) / len(amounts),
            "min_amount": min(amounts),
            "max_amount": max(amounts),
            "amount_ranges": ranges
        }