"""
Proxy Rotation Support
Manages a pool of proxies for rotating requests.
"""

import random
import itertools
from typing import Optional


class ProxyRotator:
    """Manages proxy rotation for web scraping."""
    
    def __init__(self, proxies: list[str] = None):
        """
        Initialize with a list of proxies.
        
        Args:
            proxies: List of proxy URLs. Format: "protocol://host:port" or "host:port"
                     Supported: http://, https://, socks5://
        """
        self.proxies = proxies or []
        self.proxy_cycle = itertools.cycle(self.proxies) if self.proxies else None
        self.failed_proxies = set()
        self.success_count = {}
        self.fail_count = {}
    
    @classmethod
    def from_file(cls, filepath: str) -> 'ProxyRotator':
        """Load proxies from a text file (one per line)."""
        with open(filepath, 'r') as f:
            proxies = [line.strip() for line in f if line.strip() and not line.startswith('#')]
        return cls(proxies)
    
    @classmethod
    def from_string(cls, proxy_string: str) -> 'ProxyRotator':
        """Load proxies from a comma-separated string."""
        proxies = [p.strip() for p in proxy_string.split(',') if p.strip()]
        return cls(proxies)
    
    def get_proxy(self) -> Optional[dict]:
        """Get the next proxy in rotation."""
        if not self.proxy_cycle:
            return None
        
        # Try to get a non-failed proxy
        for _ in range(len(self.proxies)):
            proxy_url = next(self.proxy_cycle)
            if proxy_url not in self.failed_proxies:
                return {'http': proxy_url, 'https': proxy_url}
        
        # If all failed, reset and try again
        self.failed_proxies.clear()
        proxy_url = next(self.proxy_cycle)
        return {'http': proxy_url, 'https': proxy_url}
    
    def get_playwright_proxy(self) -> Optional[dict]:
        """Get proxy config for Playwright."""
        if not self.proxy_cycle:
            return None
        
        for _ in range(len(self.proxies)):
            proxy_url = next(self.proxy_cycle)
            if proxy_url not in self.failed_proxies:
                # Parse proxy URL
                from urllib.parse import urlparse
                parsed = urlparse(proxy_url)
                
                proxy_config = {
                    'server': f'{parsed.scheme}://{parsed.hostname}:{parsed.port}',
                }
                
                if parsed.username:
                    proxy_config['username'] = parsed.username
                if parsed.password:
                    proxy_config['password'] = parsed.password
                
                return proxy_config
        
        # Reset and retry
        self.failed_proxies.clear()
        proxy_url = next(self.proxy_cycle)
        from urllib.parse import urlparse
        parsed = urlparse(proxy_url)
        return {
            'server': f'{parsed.scheme}://{parsed.hostname}:{parsed.port}',
            'username': parsed.username,
            'password': parsed.password,
        } if parsed.username else {'server': f'{parsed.scheme}://{parsed.hostname}:{parsed.port}'}
    
    def mark_success(self, proxy_url: str):
        """Mark a proxy as successful."""
        self.success_count[proxy_url] = self.success_count.get(proxy_url, 0) + 1
    
    def mark_failed(self, proxy_url: str):
        """Mark a proxy as failed."""
        self.fail_count[proxy_url] = self.fail_count.get(proxy_url, 0) + 1
        # After 3 failures, mark as failed
        if self.fail_count[proxy_url] >= 3:
            self.failed_proxies.add(proxy_url)
    
    @property
    def count(self) -> int:
        """Number of available proxies."""
        return len(self.proxies) - len(self.failed_proxies)
    
    @property
    def has_proxies(self) -> bool:
        """Whether any proxies are available."""
        return len(self.proxies) > 0
    
    def get_stats(self) -> dict:
        """Get proxy pool statistics."""
        return {
            'total': len(self.proxies),
            'available': self.count,
            'failed': len(self.failed_proxies),
            'success_counts': self.success_count,
            'fail_counts': self.fail_count,
        }


# Free proxy sources (for testing only - not reliable for production)
FREE_PROXY_SOURCES = [
    'https://raw.githubusercontent.com/TheSpeedX/PROXY-List/master/http.txt',
    'https://raw.githubusercontent.com/ShiftyTR/Proxy-List/master/http.txt',
    'https://raw.githubusercontent.com/monosans/proxy-list/main/proxies/http.txt',
]


def fetch_free_proxies(limit: int = 50) -> list[str]:
    """Fetch free proxies from public sources (for testing)."""
    import requests
    
    proxies = set()
    
    for source in FREE_PROXY_SOURCES:
        try:
            resp = requests.get(source, timeout=10)
            if resp.status_code == 200:
                for line in resp.text.strip().split('\n'):
                    line = line.strip()
                    if line and ':' in line:
                        proxy = f'http://{line}'
                        proxies.add(proxy)
                        if len(proxies) >= limit:
                            break
        except:
            pass
        
        if len(proxies) >= limit:
            break
    
    return list(proxies)[:limit]
