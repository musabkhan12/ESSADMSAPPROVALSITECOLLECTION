const SECRET_KEY = '2911e4410144e1f1677908c1f3932dbf4ed28a48cb69cfeaa13eec8865b0fe81';

async function getKey(): Promise<CryptoKey> {
    const enc = new TextEncoder();
    return await crypto.subtle.importKey(
        'raw',
        enc.encode(SECRET_KEY.slice(0, 32)), // must be exactly 32 chars
        { name: 'AES-GCM' },
        false,
        ['encrypt', 'decrypt']
    );
}

// ✅ Call this on BUTTON CLICK to generate token
export async function encryptParams(docId: string | number, mode: string): Promise<string> {

    const key = await getKey();

    const iv = crypto.getRandomValues(new Uint8Array(12));

    const payload = JSON.stringify({
        docId: String(docId),
        mode
    });

    const encoded = new TextEncoder().encode(payload);

    const ciphertext = await crypto.subtle.encrypt(
        { name: 'AES-GCM', iv },
        key,
        encoded
    );

    const combined = new Uint8Array(
        iv.byteLength + ciphertext.byteLength
    );

    combined.set(iv, 0);

    combined.set(
        new Uint8Array(ciphertext),
        iv.byteLength
    );

    return btoa([].slice.call(combined).map((b: number) => String.fromCharCode(b)).join('')).replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');
}

// ✅ Call this in PDFTRON WEBPART to decrypt token
export async function decryptParams(token: string): Promise<{ docId: string; mode: string } | null> {
    try {
        const key = await getKey();

        const b64 = token.replace(/-/g, '+').replace(/_/g, '/');
        const binary = atob(b64);
        const bytes = Uint8Array.from(binary, c => c.charCodeAt(0));

        const iv = bytes.slice(0, 12);
        const ciphertext = bytes.slice(12);

        const decrypted = await crypto.subtle.decrypt(
            { name: 'AES-GCM', iv },
            key,
            ciphertext
        );

        return JSON.parse(new TextDecoder().decode(decrypted));
    } catch {
        return null; // tampered or invalid
    }
}