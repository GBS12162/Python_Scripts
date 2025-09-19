"""
Servizio per l'esportazione di report in vari formati.
Supporta Excel, CSV, PDF e altri formati di output.
"""

import logging
import pandas as pd
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any, Optional
import json

from models.transaction_reporting import RejectionReport, MonthlyReportConfig, ProcessingResult
from attivita.controlli_di_linea.services.excel_service import ExcelService
from utils.date_utils import get_current_timestamp
from utils.file_utils import ensure_directory


class ReportExportService:
    """Servizio per l'esportazione di report in vari formati."""
    
    def __init__(self):
        """Inizializza il servizio."""
        self.logger = logging.getLogger(__name__)
        self.excel_service = ExcelService()
    
    def export_report(
        self, 
        report: RejectionReport, 
        config: MonthlyReportConfig,
        statistics: Optional[Dict[str, Any]] = None
    ) -> ProcessingResult:
        """
        Esporta il report nel formato specificato.
        
        Args:
            report: Report da esportare
            config: Configurazione di esportazione
            statistics: Statistiche aggiuntive (opzionale)
            
        Returns:
            Risultato dell'operazione di esportazione
        """
        try:
            start_time = datetime.now()
            self.logger.info(f"Inizio esportazione report in formato {config.export_format}")
            
            # Assicura che la directory di output esista
            output_dir = Path(config.output_directory)
            ensure_directory(str(output_dir))
            
            output_files = []
            
            if config.export_format.lower() == "excel":
                excel_file = self._export_to_excel(report, config, statistics)
                output_files.append(excel_file)
            
            elif config.export_format.lower() == "csv":
                csv_files = self._export_to_csv(report, config)
                output_files.extend(csv_files)
            
            elif config.export_format.lower() == "json":
                json_file = self._export_to_json(report, config, statistics)
                output_files.append(json_file)
            
            elif config.export_format.lower() == "all":
                # Esporta in tutti i formati
                excel_file = self._export_to_excel(report, config, statistics)
                csv_files = self._export_to_csv(report, config)
                json_file = self._export_to_json(report, config, statistics)
                
                output_files.append(excel_file)
                output_files.extend(csv_files)
                output_files.append(json_file)
            
            else:
                raise ValueError(f"Formato di esportazione non supportato: {config.export_format}")
            
            processing_time = (datetime.now() - start_time).total_seconds()
            
            return ProcessingResult(
                success=True,
                message=f"Report esportato con successo in {len(output_files)} file",
                report=report,
                output_files=output_files,
                processing_time=processing_time
            )
            
        except Exception as e:
            self.logger.error(f"Errore nell'esportazione del report: {e}")
            return ProcessingResult(
                success=False,
                message=f"Errore nell'esportazione: {str(e)}",
                errors=[str(e)]
            )
    
    def _export_to_excel(
        self, 
        report: RejectionReport, 
        config: MonthlyReportConfig,
        statistics: Optional[Dict[str, Any]] = None
    ) -> str:
        """Esporta il report in formato Excel."""
        try:
            output_file = Path(config.output_directory) / f"{config.get_filename()}.xlsx"
            
            # Prepara i dati per Excel
            transactions_data = self._prepare_transactions_data(report.transactions)
            summary_data = self._prepare_summary_data(report, statistics)
            
            # Crea il workbook Excel
            with pd.ExcelWriter(str(output_file), engine='xlsxwriter') as writer:
                
                # Foglio riepilogo
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, sheet_name='Riepilogo', index=False)
                
                # Foglio transazioni rifiutate
                if transactions_data:
                    transactions_df = pd.DataFrame(transactions_data)
                    transactions_df.to_excel(writer, sheet_name='Transazioni Rifiutate', index=False)
                
                # Foglio statistiche per tipo
                if report.rejection_by_type:
                    type_stats = pd.DataFrame([
                        {"Tipo Transazione": k, "Numero Rifiuti": v}
                        for k, v in report.rejection_by_type.items()
                    ])
                    type_stats.to_excel(writer, sheet_name='Statistiche per Tipo', index=False)
                
                # Foglio statistiche per motivo
                if report.rejection_by_reason:
                    reason_stats = pd.DataFrame([
                        {"Motivo Rifiuto": k, "Numero Occorrenze": v}
                        for k, v in report.rejection_by_reason.items()
                    ])
                    reason_stats.to_excel(writer, sheet_name='Statistiche per Motivo', index=False)
                
                # Formattazione
                self._format_excel_sheets(writer, report)
            
            self.logger.info(f"Report Excel creato: {output_file}")
            return str(output_file)
            
        except Exception as e:
            self.logger.error(f"Errore nella creazione del file Excel: {e}")
            raise
    
    def _export_to_csv(self, report: RejectionReport, config: MonthlyReportConfig) -> List[str]:
        """Esporta il report in formato CSV (multipli file)."""
        try:
            output_files = []
            base_filename = config.get_filename()
            output_dir = Path(config.output_directory)
            
            # File transazioni rifiutate
            if report.transactions:
                transactions_data = self._prepare_transactions_data(report.transactions)
                transactions_df = pd.DataFrame(transactions_data)
                transactions_file = output_dir / f"{base_filename}_transazioni.csv"
                transactions_df.to_csv(transactions_file, index=False, encoding='utf-8-sig')
                output_files.append(str(transactions_file))
            
            # File riepilogo
            summary_data = self._prepare_summary_data(report)
            summary_df = pd.DataFrame(summary_data)
            summary_file = output_dir / f"{base_filename}_riepilogo.csv"
            summary_df.to_csv(summary_file, index=False, encoding='utf-8-sig')
            output_files.append(str(summary_file))
            
            # File statistiche per tipo
            if report.rejection_by_type:
                type_stats = pd.DataFrame([
                    {"Tipo_Transazione": k, "Numero_Rifiuti": v}
                    for k, v in report.rejection_by_type.items()
                ])
                type_file = output_dir / f"{base_filename}_statistiche_tipo.csv"
                type_stats.to_csv(type_file, index=False, encoding='utf-8-sig')
                output_files.append(str(type_file))
            
            self.logger.info(f"File CSV creati: {len(output_files)} file")
            return output_files
            
        except Exception as e:
            self.logger.error(f"Errore nella creazione dei file CSV: {e}")
            raise
    
    def _export_to_json(
        self, 
        report: RejectionReport, 
        config: MonthlyReportConfig,
        statistics: Optional[Dict[str, Any]] = None
    ) -> str:
        """Esporta il report in formato JSON."""
        try:
            output_file = Path(config.output_directory) / f"{config.get_filename()}.json"
            
            # Prepara i dati JSON
            json_data = {
                "report_info": {
                    "report_id": report.report_id,
                    "generation_date": report.generation_date.isoformat(),
                    "period_start": report.period_start.isoformat(),
                    "period_end": report.period_end.isoformat()
                },
                "summary": {
                    "total_transactions": report.total_transactions,
                    "rejected_transactions": report.rejected_transactions,
                    "rejection_rate": report.rejection_rate,
                    "total_amount": float(report.total_amount),
                    "rejected_amount": float(report.rejected_amount)
                },
                "analysis": {
                    "rejection_by_type": report.rejection_by_type,
                    "rejection_by_reason": report.rejection_by_reason
                },
                "transactions": [
                    {
                        "transaction_id": t.transaction_id,
                        "account_number": t.account_number,
                        "amount": float(t.amount),
                        "currency": t.currency,
                        "transaction_date": t.transaction_date.isoformat(),
                        "transaction_type": t.transaction_type,
                        "status": t.status,
                        "rejection_reason": t.rejection_reason,
                        "rejection_code": t.rejection_code,
                        "merchant_name": t.merchant_name
                    }
                    for t in report.transactions
                ]
            }
            
            if statistics:
                json_data["statistics"] = statistics
            
            # Salva il file JSON
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(json_data, f, indent=2, ensure_ascii=False)
            
            self.logger.info(f"Report JSON creato: {output_file}")
            return str(output_file)
            
        except Exception as e:
            self.logger.error(f"Errore nella creazione del file JSON: {e}")
            raise
    
    def _prepare_transactions_data(self, transactions: List) -> List[Dict[str, Any]]:
        """Prepara i dati delle transazioni per l'esportazione."""
        return [
            {
                "ID Transazione": t.transaction_id,
                "Numero Conto": t.account_number,
                "Importo": float(t.amount),
                "Valuta": t.currency,
                "Data Transazione": t.transaction_date.strftime("%Y-%m-%d %H:%M:%S"),
                "Tipo Transazione": t.transaction_type,
                "Status": t.status,
                "Motivo Rifiuto": t.rejection_reason or "",
                "Codice Rifiuto": t.rejection_code or "",
                "Data Elaborazione": t.processing_date.strftime("%Y-%m-%d %H:%M:%S") if t.processing_date else "",
                "ID Merchant": t.merchant_id or "",
                "Nome Merchant": t.merchant_name or "",
                "Carta (Mascherata)": t.card_number_masked or ""
            }
            for t in transactions
        ]
    
    def _prepare_summary_data(
        self, 
        report: RejectionReport, 
        statistics: Optional[Dict[str, Any]] = None
    ) -> List[Dict[str, Any]]:
        """Prepara i dati del riepilogo per l'esportazione."""
        data = [
            {"Metrica": "ID Report", "Valore": report.report_id},
            {"Metrica": "Data Generazione", "Valore": report.generation_date.strftime("%Y-%m-%d %H:%M:%S")},
            {"Metrica": "Periodo Inizio", "Valore": report.period_start.strftime("%Y-%m-%d")},
            {"Metrica": "Periodo Fine", "Valore": report.period_end.strftime("%Y-%m-%d")},
            {"Metrica": "Totale Transazioni", "Valore": report.total_transactions},
            {"Metrica": "Transazioni Rifiutate", "Valore": report.rejected_transactions},
            {"Metrica": "Tasso di Rifiuto (%)", "Valore": f"{report.rejection_rate:.2f}%"},
            {"Metrica": "Importo Totale", "Valore": f"{float(report.total_amount):.2f}"},
            {"Metrica": "Importo Rifiutato", "Valore": f"{float(report.rejected_amount):.2f}"}
        ]
        
        if statistics:
            if "amounts" in statistics:
                amounts = statistics["amounts"]
                data.append({"Metrica": "Importo Medio Transazione", "Valore": f"{amounts.get('average_transaction', 0):.2f}"})
                data.append({"Metrica": "Importo Medio Rifiutato", "Valore": f"{amounts.get('average_rejected', 0):.2f}"})
        
        return data
    
    def _format_excel_sheets(self, writer, report: RejectionReport):
        """Applica formattazione ai fogli Excel."""
        try:
            workbook = writer.book
            
            # Formato per i titoli
            title_format = workbook.add_format({
                'bold': True,
                'font_size': 12,
                'bg_color': '#D7E4BC',
                'border': 1
            })
            
            # Formato per i numeri
            number_format = workbook.add_format({'num_format': '#,##0.00'})
            
            # Formato per le percentuali
            percent_format = workbook.add_format({'num_format': '0.00%'})
            
            # Applica formattazione ai fogli se esistono
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                worksheet.set_column('A:Z', 15)  # Larghezza colonne
                
        except Exception as e:
            self.logger.warning(f"Errore nella formattazione Excel: {e}")