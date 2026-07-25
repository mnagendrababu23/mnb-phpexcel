# Adapter Package Strategy

Keep the core package lightweight and framework-neutral.

Recommended optional packages:

```text
mnb/mnb-phpexcel-laravel
mnb/mnb-phpexcel-codeigniter
mnb/mnb-phpexcel-slim
mnb/mnb-phpexcel-phpspreadsheet-adapter
mnb/mnb-phpexcel-openspout-adapter
```

## Laravel adapter package idea

Do not put Laravel into the core package. Create a separate package with:

- Service provider
- Config publish command
- Queue job wrapper
- Storage disk adapter
- Validation rule bridge
- Eloquent/import target adapter
- Controller examples for upload/import/status/download failed rows

## CodeIgniter/Slim examples

Core already works in custom apps because it accepts `.env`, PHP config files, constants, arrays, and existing PDO connections.

