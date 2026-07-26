<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Domain;

use Mnb\PHPExcel\Support\MnbExcelException;

final class DomainImportPreset
{
    /**
     * @param array<string,array<string,mixed>> $fields
     * @param list<string> $uniqueBy
     * @param list<callable(array<string,mixed>,int):list<string>|string|null> $rowValidators
     */
    public function __construct(
        public readonly DomainImportType $type,
        public readonly string $defaultTable,
        public readonly string $description,
        private readonly array $fields,
        private readonly array $uniqueBy = [],
        private readonly array $rowValidators = []
    ) {
        if ($fields === []) {
            throw new MnbExcelException('A domain import preset must define at least one field.');
        }
    }

    /** @return list<string> */
    public function columns(): array { return array_keys($this->fields); }
    /** @return array<string,array<string,mixed>> */
    public function fields(): array { return $this->fields; }

    /** @return array<string,list<string>> */
    public function aliases(): array
    {
        $out = [];
        foreach ($this->fields as $name => $definition) {
            $values = array_merge([$name], array_map('strval', (array) ($definition['aliases'] ?? [])));
            $out[$name] = array_values(array_unique(array_filter(array_map('trim', $values), static fn(string $v): bool => $v !== '')));
        }
        return $out;
    }

    /** @return array<string,string> */
    public function rules(): array
    {
        $out = [];
        foreach ($this->fields as $name => $definition) {
            $rule = trim((string) ($definition['rule'] ?? ''));
            if ($rule !== '') $out[$name] = $rule;
        }
        return $out;
    }

    /** @return array<string,mixed> */
    public function defaults(): array
    {
        $out = [];
        foreach ($this->fields as $name => $definition) {
            if (array_key_exists('default', $definition)) $out[$name] = $definition['default'];
        }
        return $out;
    }

    /** @return list<string> */
    public function requiredColumns(): array
    {
        $out = [];
        foreach ($this->fields as $name => $definition) {
            $rules = explode('|', (string) ($definition['rule'] ?? ''));
            if (($definition['required'] ?? false) === true || in_array('required', $rules, true)) $out[] = $name;
        }
        return $out;
    }

    /** @return list<string> */
    public function uniqueBy(): array { return $this->uniqueBy; }
    /** @return list<callable(array<string,mixed>,int):list<string>|string|null> */
    public function rowValidators(): array { return $this->rowValidators; }

    /** @return array<int,array<string,mixed>> */
    public function templateColumns(): array
    {
        $columns = [];
        foreach ($this->fields as $name => $definition) {
            $rule = (string) ($definition['rule'] ?? '');
            $column = [
                'name' => $name,
                'header' => (string) ($definition['header'] ?? ucwords(str_replace('_', ' ', $name))),
                'description' => (string) ($definition['description'] ?? ''),
                'example' => $definition['example'] ?? '',
                'required' => ($definition['required'] ?? false) === true || in_array('required', explode('|', $rule), true),
            ];
            if (isset($definition['list']) && is_array($definition['list'])) $column['list'] = array_values($definition['list']);
            if (isset($definition['validation']) && is_array($definition['validation'])) $column['validation'] = $definition['validation'];
            $columns[] = $column;
        }
        return $columns;
    }

    /** @return array<string,mixed> */
    public function toArray(): array
    {
        return [
            'domain' => $this->type->value,
            'default_table' => $this->defaultTable,
            'description' => $this->description,
            'columns' => $this->columns(),
            'required_columns' => $this->requiredColumns(),
            'unique_by' => $this->uniqueBy,
            'aliases' => $this->aliases(),
            'rules' => $this->rules(),
            'defaults' => $this->defaults(),
            'fields' => $this->fields,
        ];
    }
}
