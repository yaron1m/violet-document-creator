# Violet Document Creator

Offer PDF creator for [Violet](https://github.com/yaron1m/violet).

## Installation
Start cmd as Administrator, and run
```sh
VioletDocumentCreator.exe install
```
This will create a new URI Scheme for violet and point it to the location of VioletDocumentCreator.exe.

To uninstall:
```sh
VioletDocumentCreator.exe uninstall
```

If you need to move this file, uninstall and install again in the new location.

## Usage Example
```js
violet:{
  "topic": [
    "ISO-14001",
    "IS0-9001"
  ],
  "email": "email@mail.com",
  "organizationName": "orgName",
  "contactFirstName": "firstName",
  "contactLastName": "lastName",
  "id": "5000",
  "orderCreationDate": "2017-05-01T07:34:42-5:00",
  "contactPhone1": "052-2222222",
  "contactPhone2": "03-3333333"
}
```

## Author

**Yaron Malin** - [yaron1m](https://github.com/yaron1m)

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details.
