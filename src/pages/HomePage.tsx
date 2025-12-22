import { useEffect, useRef, useState } from 'react';
import { Html5Qrcode, Html5QrcodeScannerState } from 'html5-qrcode';
import { AppLayout } from '@/components/layout/AppLayout';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Camera, CameraOff, Zap } from 'lucide-react';
import { Toaster, toast } from 'sonner';
import { useInventoryStore } from '@/hooks/use-inventory';
import { useAuthStore } from '@/hooks/use-auth';
import { cn, playBeep } from '@/lib/utils';
import type { ScanEvent } from '@shared/types';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { useTranslation } from 'react-i18next';
const QR_READER_ID = 'qr-reader';
export function HomePage() {
  const { t } = useTranslation();
  const SCAN_EVENTS: { value: ScanEvent; label: string }[] = [
    { value: 'PRODUCTION_SCAN', label: t('scanner.events.PRODUCTION_SCAN') },
    { value: 'MIRAZUR_BOOK_SCAN', label: t('scanner.events.MIRAZUR') },
    { value: 'OFFICE_BOOK_SCAN', label: t('scanner.events.OFFICE SCAN') },
    
  ];
  const [scanStatus, setScanStatus] = useState<'idle' | 'scanning' | 'success' | 'error' | 'permission_denied'>('idle');
  const [lastScan, setLastScan] = useState<string | null>(null);
  const [showSuccessOverlay, setShowSuccessOverlay] = useState(false);
  const [selectedEvent, setSelectedEvent] = useState<ScanEvent | null>(null);
  const [selectedB2BClient, setSelectedB2BClient] = useState<string | null>(null);
  const html5QrCodeRef = useRef<Html5Qrcode | null>(null);
  const addScan = useInventoryStore((state) => state.addScan);
  const b2bClients = useInventoryStore((state) => state.b2bClients);
  const fetchB2BClients = useInventoryStore((state) => state.fetchB2BClients);
  const currentUser = useAuthStore((state) => state.currentUser);
  const cooldownRef = useRef(false);
  useEffect(() => {
    fetchB2BClients();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  const onScanSuccess = (decodedText: string) => {
    if (cooldownRef.current || !currentUser?.location || !selectedEvent) return;
    if (selectedEvent === 'DELIVERY_B2B' && !selectedB2BClient) return;
    cooldownRef.current = true;
    setScanStatus('success');
    setLastScan(decodedText);
    playBeep();
    if (navigator.vibrate) navigator.vibrate(200);
    addScan(decodedText, selectedEvent, currentUser.location, selectedB2BClient || undefined);
    setShowSuccessOverlay(true);
    setTimeout(() => setShowSuccessOverlay(false), 500);
    setTimeout(() => {
      cooldownRef.current = false;
      if (html5QrCodeRef.current?.isScanning) {
        setScanStatus('scanning');
      }
    }, 2000);
  };
  const stopScanner = async () => {
    if (html5QrCodeRef.current && html5QrCodeRef.current.isScanning) {
      try {
        await html5QrCodeRef.current.stop();
      } catch (err) {
        console.error('Failed to stop scanner:', err);
      }
    }
    setScanStatus('idle');
  };
  const startScanner = async () => {
    if (!html5QrCodeRef.current) {
      html5QrCodeRef.current = new Html5Qrcode(QR_READER_ID, { verbose: false });
    }
    const qrCode = html5QrCodeRef.current;
    if (qrCode.getState() === Html5QrcodeScannerState.SCANNING) return;
    setScanStatus('scanning');
    setLastScan(null);
    try {
      await qrCode.start(
        { facingMode: 'environment' },
        { fps: 10, qrbox: { width: 250, height: 250 } },
        onScanSuccess,
        (errorMessage) => { /* ignore */ }
      );
    } catch (err) {
      console.error('Camera permission error:', err);
      setScanStatus('permission_denied');
      toast.error(t('scanner.status.permissionDenied'));
    }
  };
  useEffect(() => {
    return () => {
      stopScanner();
    };
  }, []);
  const handleEventChange = (value: string) => {
    const newEvent = value as ScanEvent;
    setSelectedEvent(newEvent);
    setSelectedB2BClient(null);
    if (scanStatus === 'scanning') {
      stopScanner();
      toast.info(t('scanner.toast.eventChange'));
    }
  };
  const getStatusMessage = () => {
    switch (scanStatus) {
      case 'scanning': return t('scanner.status.scanning');
      case 'success': return t('scanner.status.success', { lastScan });
      case 'permission_denied': return t('scanner.status.permissionDenied');
      case 'error': return t('scanner.status.error');
      default: return t('scanner.status.idle');
    }
  };
  const isStartDisabled = !selectedEvent || (selectedEvent === 'DELIVERY_B2B' && !selectedB2BClient);
  return (
    <AppLayout>
      <div className="min-h-screen flex flex-col items-center justify-center bg-gray-50 dark:bg-gray-900 p-4">
        <Card className="w-full max-w-md shadow-lg animate-fade-in">
          <CardHeader>
            <CardTitle className="text-center text-2xl font-bold text-foreground">
              {t('scanner.title')}
            </CardTitle>
          </CardHeader>
          <CardContent className="space-y-6">
            <div className={cn("relative aspect-square w-full rounded-lg overflow-hidden bg-secondary border-dashed border-2 border-border flex items-center justify-center transition-all duration-300", { "border-green-500 shadow-lg shadow-green-500/20": showSuccessOverlay })}>
              <div id={QR_READER_ID} className="w-full h-full" />
              {scanStatus !== 'scanning' && (
                <div className="absolute inset-0 flex flex-col items-center justify-center bg-secondary/80 backdrop-blur-sm">
                  <Camera className="w-16 h-16 text-muted-foreground" />
                  <p className="mt-2 text-muted-foreground">{t('scanner.cameraOff')}</p>
                </div>
              )}
              {showSuccessOverlay && (
                <div className="absolute inset-0 bg-green-500/20 pointer-events-none" />
              )}
            </div>
            <div className="space-y-4">
              <div>
                <label className="text-sm font-medium text-muted-foreground">{t('scanner.scanEvent')}</label>
                <Select onValueChange={handleEventChange} value={selectedEvent || ''}>
                  <SelectTrigger><SelectValue placeholder={t('scanner.selectEvent')} /></SelectTrigger>
                  <SelectContent>
                    {SCAN_EVENTS.map(event => (<SelectItem key={event.value} value={event.value}>{event.label}</SelectItem>))}
                  </SelectContent>
                </Select>
              </div>
              {selectedEvent === 'DELIVERY_B2B' && (
                <div className="animate-fade-in">
                  <label className="text-sm font-medium text-muted-foreground">{t('scanner.b2bClient')}</label>
                  <Select onValueChange={(value) => setSelectedB2BClient(value)} value={selectedB2BClient || ''}>
                    <SelectTrigger><SelectValue placeholder={t('scanner.selectClient')} /></SelectTrigger>
                    <SelectContent>
                      {b2bClients.map(client => (<SelectItem key={client.clientId} value={client.clientId}>{client.clientName}</SelectItem>))}
                    </SelectContent>
                  </Select>
                </div>
              )}
            </div>
            <div className="text-center p-3 rounded-lg bg-muted text-muted-foreground font-medium transition-colors duration-300">
              {getStatusMessage()}
            </div>
            <div className="flex gap-4">
              {scanStatus === 'scanning' ? (
                <Button onClick={stopScanner} variant="destructive" className="w-full">
                  <CameraOff className="mr-2 h-4 w-4" /> {t('scanner.stop')}
                </Button>
              ) : (
                <Button onClick={startScanner} className="w-full" disabled={isStartDisabled}>
                  <Zap className="mr-2 h-4 w-4" /> {t('scanner.start')}
                </Button>
              )}
            </div>
          </CardContent>
        </Card>
        <footer className="absolute bottom-4 text-center text-muted-foreground/80 text-sm">
          <p>Built with ❤️ at Cloudflare</p>
        </footer>
      </div>
      <Toaster richColors closeButton />
    </AppLayout>
  );
}
