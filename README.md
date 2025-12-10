<a id="readme-top"></a>

[![Contributors][contributors-shield]][contributors-url]
[![Forks][forks-shield]][forks-url]
[![Stargazers][stars-shield]][stars-url]
[![Issues][issues-shield]][issues-url]
[![MIT License][license-shield]][license-url]
[![LinkedIn][linkedin-shield]][linkedin-url]

<!-- PROJECT LOGO -->
<br />
<div align="center">

<h3 align="center">Google Doc Sheet Bypass</h3>

  <p align="center">
    Bypass Docs and Sheets print and download permission
    <br />
    <a href="https://github.com/fathulfahmy/google-doc-sheet-bypass/issues/new?labels=bug&template=bug-report---.md">Report Bug</a>
    &middot;
    <a href="https://github.com/fathulfahmy/google-doc-sheet-bypass/issues/new?labels=enhancement&template=feature-request---.md">Request Feature</a>
  </p>
</div>

<!-- TABLE OF CONTENTS -->
<details>
  <summary>Table of Contents</summary>
  <ol>
    <li>
      <a href="#about-the-project">About The Project</a>
      <ul>
        <li><a href="#built-with">Built With</a></li>
      </ul>
    </li>
    <li>
      <a href="#getting-started">Getting Started</a>
      <ul>
        <li><a href="#installation">Installation</a></li>
        <li><a href="#usage">Usage</a></li>
      </ul>
    </li>
    <li><a href="#usage">Usage</a></li>
    <li><a href="#contributing">Contributing</a></li>
    <li><a href="#license">License</a></li>
    <li><a href="#contact">Contact</a></li>
    <li><a href="#acknowledgments">Acknowledgments</a></li>
  </ol>
</details>

<!-- ABOUT THE PROJECT -->

## About The Project

![Product Name Screen Shot][product-screenshot]

<p align="right">(<a href="#readme-top">back to top</a>)</p>

### Built With

- Python
- [BeautifulSoup](https://beautiful-soup-4.readthedocs.io/en/latest/)
- [python-docx](https://pypi.org/project/python-docx/)

<p align="right">(<a href="#readme-top">back to top</a>)</p>

<!-- GETTING STARTED -->

## Getting Started

## Installation

1. Clone repo

   ```sh
   git clone https://github.com/fathulfahmy/google-doc-sheet-bypass.git
   ```
2. Navigate to project directory
    ```sh
    cd google-doc-sheet-bypass
    ```

## Usage

### Option 1: Python

**Prerequisites**

- [Python 3.12](https://www.python.org/downloads/)
- [UV](https://docs.astral.sh/uv/getting-started/installation/)

1. Run app

```sh
uv sync
uv run uvicorn main:app
```

2. Open browser http://localhost:8000

## Option 2: Docker

**Prerequisites**

- [Docker](https://docs.docker.com/engine/install/)

1. Run app

   ```sh
   docker compose up
   ```

2. Open browser http://localhost:8000

> [!NOTE]  
> Rebuild docker container after making changes to the code
>
> ```sh
> docker compose up --build --force-recreate
> ```

### Development

1. Install dependencies

   ```sh
   uv sync
   uv run pre-commit install
   ```

2. Run project

   ```sh
   uv run uvicorn main:app
   ```

<p align="right">(<a href="#readme-top">back to top</a>)</p>

<!-- CONTRIBUTING -->

## Contributing

Contributions are what make the open source community such an amazing place to learn, inspire, and create. Any contributions you make are **greatly appreciated**.

If you have a suggestion that would make this better, please fork the repo and create a pull request. You can also simply open an issue with the tag "enhancement".
Don't forget to give the project a star! Thanks again!

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feat/amazing-feature`)
3. Commit your Changes (`git commit -m 'feat: add some amazing feature'`)
4. Push to the Branch (`git push origin feat/amazing-feature`)
5. Open a Pull Request

<p align="right">(<a href="#readme-top">back to top</a>)</p>

### Top contributors:

<a href="https://github.com/fathulfahmy/google-doc-sheet-bypass/graphs/contributors">
  <img src="https://contrib.rocks/image?repo=fathulfahmy/google-doc-sheet-bypass" alt="contrib.rocks image" />
</a>

<!-- LICENSE -->

## License

Distributed under the MIT License. See `LICENSE` for more information.

<p align="right">(<a href="#readme-top">back to top</a>)</p>

<!-- CONTACT -->

## Contact

Your Name - [@fathulfahmy](https://linkedin.com/in/fathulfahmy) - fathulfahmy@protonmail.com

Project Link: [https://github.com/fathulfahmy/google-doc-sheet-bypass](https://github.com/fathulfahmy/google-doc-sheet-bypass)

<p align="right">(<a href="#readme-top">back to top</a>)</p>

<!-- ACKNOWLEDGMENTS -->

## Acknowledgments

- [bypass-google](https://github.com/EdwardX29/bypass-google)

<p align="right">(<a href="#readme-top">back to top</a>)</p>

<!-- MARKDOWN LINKS & IMAGES -->
<!-- https://www.markdownguide.org/basic-syntax/#reference-style-links -->

[contributors-shield]: https://img.shields.io/github/contributors/fathulfahmy/google-doc-sheet-bypass.svg?style=for-the-badge
[contributors-url]: https://github.com/fathulfahmy/google-doc-sheet-bypass/graphs/contributors
[forks-shield]: https://img.shields.io/github/forks/fathulfahmy/google-doc-sheet-bypass.svg?style=for-the-badge
[forks-url]: https://github.com/fathulfahmy/google-doc-sheet-bypass/network/members
[stars-shield]: https://img.shields.io/github/stars/fathulfahmy/google-doc-sheet-bypass.svg?style=for-the-badge
[stars-url]: https://github.com/fathulfahmy/google-doc-sheet-bypass/stargazers
[issues-shield]: https://img.shields.io/github/issues/fathulfahmy/google-doc-sheet-bypass.svg?style=for-the-badge
[issues-url]: https://github.com/fathulfahmy/google-doc-sheet-bypass/issues
[license-shield]: https://img.shields.io/github/license/fathulfahmy/google-doc-sheet-bypass.svg?style=for-the-badge
[license-url]: https://github.com/fathulfahmy/google-doc-sheet-bypass/blob/master/LICENSE
[linkedin-shield]: https://img.shields.io/badge/-LinkedIn-black.svg?style=for-the-badge&logo=linkedin&colorB=555
[linkedin-url]: https://linkedin.com/in/fathulfahmy
[product-screenshot]: static/screenshot.png
