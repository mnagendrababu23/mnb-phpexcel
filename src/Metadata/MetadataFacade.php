<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Metadata;

use Mnb\PHPExcel\MnbExcel;

/**
 * Developer-friendly metadata query/editor.
 * Reading is lazy: no scan occurs until read()/toArray() or a typed getter is called.
 */
final class MetadataFacade
{
    /** @var array<string,array<string,mixed>> */
    private static array $memoryCache = [];
    /** @var list<string>|null */
    private ?array $sections = null;
    /** @var array<string,mixed> */
    private array $options = [];
    /** @var array<string,mixed> */
    private array $pendingChanges = [];
    private ?MetadataResult $result = null;

    /** @param array<string,mixed> $options */
    public function __construct(private readonly string $path, array $options = [])
    {
        $this->options = $options;
    }

    public function quick(): self { return $this->profile(MetadataProfile::QUICK); }
    public function standard(): self { return $this->profile(MetadataProfile::STANDARD); }
    public function full(): self { return $this->profile(MetadataProfile::FULL); }
    public function forensic(): self { return $this->profile(MetadataProfile::FORENSIC); }

    public function profile(string $profile): self
    {
        $clone = clone $this;
        $clone->options['profile'] = MetadataProfile::normalize($profile);
        $clone->result = null;
        return $clone;
    }

    /** @param list<string> $sections */
    public function only(array $sections): self
    {
        $valid = array_values(array_unique(array_map('strval', $sections)));
        foreach ($valid as $section) {
            if (!in_array($section, MetadataReport::SECTIONS, true)) {
                throw new \InvalidArgumentException('Unknown metadata section: ' . $section);
            }
        }
        $clone = clone $this;
        $clone->sections = $valid;
        $clone->options['only_sections'] = $valid;
        $clone->result = null;
        return $clone;
    }

    public function withHash(bool $enabled = true): self
    {
        $clone = clone $this;
        $clone->options['include_hash'] = $enabled;
        $clone->result = null;
        return $clone;
    }

    public function cache(bool $enabled = true): self
    {
        $clone = clone $this;
        $clone->options['metadata_cache'] = $enabled;
        $clone->result = null;
        return $clone;
    }

    public function read(): MetadataResult
    {
        if ($this->result !== null) return $this->result;

        $cacheEnabled = (bool) ($this->options['metadata_cache'] ?? true);
        $key = $this->fingerprint();
        if ($cacheEnabled && isset(self::$memoryCache[$key])) {
            return $this->result = new MetadataResult(self::$memoryCache[$key]);
        }

        $data = MnbExcel::metaInfo($this->path, $this->options);
        if ($this->sections !== null) {
            $keep = array_flip(array_merge([
                'schema_version','status','profile','format','format_variant','mime_type',
                'capabilities','warnings','errors'
            ], $this->sections));
            $data = array_intersect_key($data, $keep);
        }

        if ($cacheEnabled) self::$memoryCache[$key] = $data;
        return $this->result = new MetadataResult($data);
    }

    /** @return array<string,mixed> */
    public function toArray(): array { return $this->read()->toArray(); }
    public function format(): string { return $this->read()->format(); }
    public function sheetCount(): int { return $this->read()->sheetCount(); }
    public function title(): ?string { return $this->read()->title(); }
    public function author(): ?string { return $this->read()->author(); }
    public function company(): ?string { return $this->read()->company(); }
    /** @return list<array<string,mixed>> */
    public function sheets(): array { return $this->read()->sheets(); }

    /** @param array<string,mixed> $changes */
    public function update(array $changes): self
    {
        $clone = clone $this;
        $clone->pendingChanges = array_replace($clone->pendingChanges, $changes);
        return $clone;
    }

    public function save(?string $destination = null): string
    {
        $destination ??= $this->path;
        if ($this->pendingChanges === []) {
            throw new \LogicException('No metadata changes were supplied. Call update() first.');
        }
        MnbExcel::updateMetaInfo($this->path, $destination, $this->pendingChanges, $this->options);
        self::clearCache();
        return $destination;
    }

    public function removePersonalInfo(?string $destination = null): string
    {
        $destination ??= $this->path;
        MnbExcel::removePersonalInfo($this->path, $destination, $this->options);
        self::clearCache();
        return $destination;
    }

    /** @return array<string,mixed> */
    public function diff(string|MetadataResult|self $other): array
    {
        $left = $this->read()->toArray();
        $right = match (true) {
            $other instanceof self => $other->read()->toArray(),
            $other instanceof MetadataResult => $other->toArray(),
            default => (new self($other, $this->options))->read()->toArray(),
        };
        return MetadataDiff::between($left, $right);
    }

    public function fingerprint(bool $withHash = false): string
    {
        $stat = @stat($this->path) ?: [];
        $parts = [
            realpath($this->path) ?: $this->path,
            (string) ($stat['size'] ?? -1),
            (string) ($stat['mtime'] ?? -1),
            json_encode($this->options, JSON_UNESCAPED_SLASHES) ?: '{}',
            json_encode($this->sections, JSON_UNESCAPED_SLASHES) ?: 'null',
        ];
        if ($withHash && is_file($this->path)) $parts[] = hash_file('sha256', $this->path) ?: '';
        return hash('sha256', implode("\0", $parts));
    }

    public static function clearCache(): void { self::$memoryCache = []; }
}
