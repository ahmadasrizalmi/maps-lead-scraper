"""
Google Sheets Integration
Export scraped data directly to Google Sheets.
"""

import json
from typing import Optional


class GoogleSheetsExporter:
    """Export data to Google Sheets using the Google Workspace skill."""
    
    def __init__(self):
        self.available = False
        try:
            import subprocess
            result = subprocess.run(
                ['python3', '-c', 'import google.oauth2; print("ok")'],
                capture_output=True, text=True
            )
            self.available = result.returncode == 0
        except:
            pass
    
    def create_spreadsheet(self, title: str, data: list[dict], query: str = '') -> Optional[str]:
        """
        Create a Google Spreadsheet with the scraped data.
        
        Returns: URL of the created spreadsheet, or None if failed.
        """
        try:
            from google.oauth2.credentials import Credentials
            from googleapiclient.discovery import build
            
            # Load credentials
            creds = self._load_credentials()
            if not creds:
                return None
            
            service = build('sheets', 'v4', credentials=creds)
            
            # Create spreadsheet
            spreadsheet = {
                'properties': {'title': title},
                'sheets': [{
                    'properties': {
                        'title': 'Business Leads',
                        'gridProperties': {'frozenRowCount': 1}
                    }
                }]
            }
            
            result = service.spreadsheets().create(body=spreadsheet).execute()
            spreadsheet_id = result['spreadsheetId']
            spreadsheet_url = result['spreadsheetUrl']
            
            # Prepare headers
            headers = [
                'No', 'Business Name', 'Category', 'Address', 'Phone',
                'Website', 'Email', 'Instagram', 'Facebook', 'TikTok',
                'Rating', 'Reviews', 'Status', 'Hours', 'Google Maps URL',
                'Description', 'Outreach Status', 'Notes'
            ]
            
            # Prepare rows
            rows = [headers]
            for i, item in enumerate(data, 1):
                social = item.get('social_media', {})
                row = [
                    str(i),
                    item.get('name', ''),
                    item.get('category', ''),
                    item.get('address', ''),
                    item.get('phone', ''),
                    item.get('website', ''),
                    item.get('email', ''),
                    social.get('instagram', item.get('social_instagram', '')),
                    social.get('facebook', item.get('social_facebook', '')),
                    social.get('tiktok', item.get('social_tiktok', '')),
                    str(item.get('rating', '')),
                    str(item.get('reviews', '')),
                    item.get('status', ''),
                    item.get('hours', ''),
                    item.get('maps_url', ''),
                    item.get('description', ''),
                    '',  # Outreach Status
                    '',  # Notes
                ]
                rows.append(row)
            
            # Write data
            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range='Business Leads!A1',
                valueInputOption='RAW',
                body={'values': rows}
            ).execute()
            
            # Format header row
            requests_body = [
                {
                    'repeatCell': {
                        'range': {'sheetId': 0, 'startRowIndex': 0, 'endRowIndex': 1},
                        'cell': {
                            'userEnteredFormat': {
                                'backgroundColor': {'red': 0.39, 'green': 0.4, 'blue': 0.95},
                                'textFormat': {
                                    'foregroundColor': {'red': 1, 'green': 1, 'blue': 1},
                                    'bold': True,
                                    'fontSize': 11
                                }
                            }
                        },
                        'fields': 'userEnteredFormat(backgroundColor,textFormat)'
                    }
                },
                {
                    'autoResizeDimensions': {
                        'dimensions': {
                            'sheetId': 0,
                            'dimension': 'COLUMNS',
                            'startIndex': 0,
                            'endIndex': len(headers)
                        }
                    }
                }
            ]
            
            service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={'requests': requests_body}
            ).execute()
            
            return spreadsheet_url
            
        except Exception as e:
            print(f"Google Sheets error: {e}")
            return None
    
    def _load_credentials(self):
        """Load Google OAuth credentials."""
        from pathlib import Path
        
        # Try common token locations
        token_paths = [
            Path.home() / '.hermes' / 'google_token.json',
            Path.home() / '.config' / 'gws' / 'token.json',
            Path.home() / '.credentials' / 'google_token.json',
        ]
        
        for path in token_paths:
            if path.exists():
                try:
                    from google.oauth2.credentials import Credentials
                    creds = Credentials.from_authorized_user_file(str(path))
                    if creds and creds.valid:
                        return creds
                    if creds and creds.expired and creds.refresh_token:
                        from google.auth.transport.requests import Request
                        creds.refresh(Request())
                        return creds
                except:
                    pass
        
        return None
    
    @property
    def is_available(self) -> bool:
        """Check if Google Sheets integration is available."""
        return self.available
