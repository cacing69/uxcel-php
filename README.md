# Uxcel PHP

Uxcel is a PHP library that can protect/unprotect some supported document for you.

uxcel-php is heavily inspired by Uxcel on sourceforge [UXCEL](https://sourceforge.net/projects/uxcel/), where is written on [Java](https://github.com/COEM/UXCEL).

Uxcel requires PHP >= 5.*

# Table of Contents

- [Installation](#installation)
- [Usage](#usage)
    - [Unprotect Excel](#unprotect-excel)
- [License](#license)

## Installation

```sh
composer require cacing69/uxcel
```

## Usage

### Autoloading
uxcel-php supports both `PSR-0` as `PSR-4` autoloaders.
```php
<?php
# When installed via composer
require_once 'vendor/autoload.php';
```

You can also load `Uxcel` shipped `PSR-0` autoloader
```php
<?php
# Load Uxcel own autoloader
require_once '/path/to/Uxcel/src/autoload.php';
```

*alternatively, you can use any another PSR-4 compliant autoloader*

### `Unprotect Excel`
Use `Uxcel\Factory` to create and initialize a uxcel instance.
```php
<?php
// use the factory to create a Uxcel instance
$uxcel = new Uxcel\Factory();

// set target file to unprotect
$target = "/path/to/file.xlsx";
$uxcel->setTarget($target);

// set destination to save result
$destination = "/path/to/destination";
$uxcel->setDestination($destination);

// lets process unprotect excel
$unprotect = $uxcel->unProtect(); // it will return uxcel instance
```

## License

Uxcel-PHP is released under the MIT License. See the bundled LICENSE file for details.
