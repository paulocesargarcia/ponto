<?php

namespace App\Security;

class RateLimiter
{
    private string $storageDir;
    private int $limit;
    private int $window;

    public function __construct(string $storageDir = 'var/ratelimit/', int $limit = 10, int $window = 600)
    {
        $this->storageDir = rtrim($storageDir, '/') . '/';
        $this->limit = $limit;
        $this->window = $window;
    }

    public function check(string $ip): bool
    {
        $hash = hash('sha256', $ip);
        $file = $this->storageDir . $hash;

        $data = ['count' => 0, 'start' => time()];

        if (file_exists($file)) {
            $content = file_get_contents($file);
            $decoded = json_decode($content, true);
            if ($decoded) {
                $data = $decoded;
            }
        }

        if (time() - $data['start'] > $this->window) {
            $data = ['count' => 1, 'start' => time()];
        } else {
            $data['count']++;
        }

        file_put_contents($file, json_encode($data));

        return $data['count'] <= $this->limit;
    }
}
