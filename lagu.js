/**
 * MUSIC PLAYER AUTO-LOOP & RANDOM START
 * Fitur: 
 * 1. Mulai acak saat refresh.
 * 2. Putar berurutan (1 -> 2 -> 3...).
 * 3. Loop kembali ke awal setelah lagu terakhir.
 * 4. Tombol kontrol melayang (Floating Button).
 * 5. Bypass Autoplay Policy (Menunggu interaksi jika diblokir).
 */

(function() {
    // KONFIGURASI LAGU
    const playlist = [
        'lagu1.mp3', 'lagu2.mp3', 'lagu3.mp3', 'lagu4.mp3', 'lagu5.mp3',
        'lagu6.mp3', 'lagu7.mp3', 'lagu8.mp3', 'lagu9.mp3', 'lagu10.mp3'
    ];

    // Mulai dari index acak (0 sampai 9)
    let currentIndex = Math.floor(Math.random() * playlist.length);
    
    // Inisialisasi Audio
    const audio = new Audio();
    audio.volume = 0.5; // Volume awal 50%
    let isPlaying = false;

    // Membuat Tombol Kontrol Melayang (UI)
    function createFloatingButton() {
        const btn = document.createElement('button');
        btn.id = 'musicPlayerBtn';
        
        // Styling CSS langsung di JS agar praktis
        Object.assign(btn.style, {
            position: 'fixed',
            bottom: '20px',
            left: '20px', // Posisi di kiri bawah agar tidak menabrak tombol lain
            zIndex: '9999',
            width: '45px',
            height: '45px',
            borderRadius: '50%',
            backgroundColor: '#db2777', // Warna Pink sesuai tema klinik
            color: 'white',
            border: '2px solid white',
            boxShadow: '0 4px 6px rgba(0,0,0,0.2)',
            cursor: 'pointer',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            fontSize: '18px',
            transition: 'all 0.3s ease'
        });

        // Icon awal (menggunakan FontAwesome yang sudah ada di index.html)
        btn.innerHTML = '<i class="fas fa-music"></i>';
        
        // Event Klik Tombol
        btn.onclick = function() {
            if (audio.paused) {
                audio.play();
                isPlaying = true;
            } else {
                audio.pause();
                isPlaying = false;
            }
            updateButtonUI();
        };

        document.body.appendChild(btn);
    }

    // Update Tampilan Tombol
    function updateButtonUI() {
        const btn = document.getElementById('musicPlayerBtn');
        if (!btn) return;

        if (!audio.paused) {
            btn.innerHTML = '<i class="fas fa-volume-up"></i>';
            btn.style.transform = 'scale(1.1)';
            btn.style.backgroundColor = '#be185d'; // Pink lebih gelap saat aktif
            // Efek denyut
            btn.style.animation = 'pulse-music 1.5s infinite';
        } else {
            btn.innerHTML = '<i class="fas fa-volume-mute"></i>';
            btn.style.transform = 'scale(1)';
            btn.style.backgroundColor = '#db2777';
            btn.style.animation = 'none';
        }
    }

    // Menambahkan animasi keyframes untuk tombol
    const styleSheet = document.createElement("style");
    styleSheet.innerText = `
        @keyframes pulse-music {
            0% { box-shadow: 0 0 0 0 rgba(219, 39, 119, 0.7); }
            70% { box-shadow: 0 0 0 10px rgba(219, 39, 119, 0); }
            100% { box-shadow: 0 0 0 0 rgba(219, 39, 119, 0); }
        }
    `;
    document.head.appendChild(styleSheet);

    // Fungsi Memutar Lagu
    function playTrack(index) {
        // Pastikan index valid (looping)
        if (index >= playlist.length) index = 0;
        
        // Set source lagu
        audio.src = playlist[index];
        audio.load();

        // Coba putar (Handle Autoplay Policy)
        const playPromise = audio.play();

        if (playPromise !== undefined) {
            playPromise.then(_ => {
                // Berhasil putar otomatis
                isPlaying = true;
                updateButtonUI();
                console.log(`Memutar: ${playlist[index]}`);
            })
            .catch(error => {
                // Autoplay dicegah browser
                console.warn("Autoplay ditahan browser. Menunggu interaksi user...");
                isPlaying = false;
                updateButtonUI();
                
                // Tambahkan listener sekali pakai: Klik dimanapun akan memulai musik
                document.addEventListener('click', function startOnInteraction() {
                    audio.play();
                    isPlaying = true;
                    updateButtonUI();
                    document.removeEventListener('click', startOnInteraction);
                }, { once: true });
            });
        }
    }

    // Event Listener: Saat lagu habis, lanjut ke lagu berikutnya
    audio.addEventListener('ended', function() {
        currentIndex++;
        if (currentIndex >= playlist.length) {
            currentIndex = 0; // Loop kembali ke lagu 1
        }
        playTrack(currentIndex);
    });

    // Jalankan saat halaman siap
    window.addEventListener('DOMContentLoaded', () => {
        createFloatingButton();
        playTrack(currentIndex); // Mulai mainkan
    });

})();
