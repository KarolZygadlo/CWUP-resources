### 1. Docker

Docker to narzędzie, które zrewolucjonizowało branżę IT od swojego wprowadzenia na rynek w 2013 roku przez Solomona Hykesa, założyciela firmy dotCloud. Zaprezentowany jako przyszłość kontenerów Linux podczas krótkiej, ale przełomowej prezentacji, Docker Hub, oferuje otwarte środowisko dla developerów, testerów i administratorów oprogramowania.

Znany z lekkiej i przenośnej konteneryzacji aplikacji i ich zależności, Docker umożliwia łatwe i bezpieczne uruchamianie oprogramowania na dowolnym serwerze Linux. To podejście zyskało uznanie i popularność wśród milionów użytkowników i licznych klientów biznesowych, oferując efektywne i bezpieczne rozwiązanie dla dostarczania produktów i usług.

W dziedzinie IT, wprowadzenie Dockera jest często postrzegane jako przełom, znacząco upraszczający prototypowanie, budowanie i wdrażanie aplikacji w różnorodnych środowiskach dzięki wykorzystaniu konteneryzacji.

### 2. Wirtualizacja a konteneryzacja

Aby zrozumieć istotę i mechanizm działania Dockera, kluczowe jest zaznajomienie się z różnicą między tradycyjną wirtualizacją a konteneryzacją. W kontekście klasycznej wirtualizacji, każda maszyna wirtualna działa na pełnym systemie operacyjnym, który z kolei jest uruchomiony na systemie operacyjnym hosta. Głównym atutem tego podejścia jest zdolność do uruchamiania licznych maszyn wirtualnych, które mogą być wyposażone w różnorodne systemy operacyjne, wszystko to na jednym hoście.

Chociaż wykorzystanie izolacji dostarcza wysokiego poziomu bezpieczeństwa dla każdej maszyny wirtualnej, tradycyjna wirtualizacja niesie ze sobą obciążenie w postaci konsumpcji znacznych zasobów oraz potrzeby przeprowadzania kompletnych instalacji i konfiguracji systemów operacyjnych. Duże zapotrzebowanie na zasoby oznacza również, że na jednym hoście może funkcjonować jedynie ograniczona liczba maszyn wirtualnych.

Konteneryzacja, będąca kluczowym elementem działania Dockera, operuje w obszarze, który określamy mianem kontenera Dockera, różniąc się tym od konwencjonalnej maszyny wirtualnej. W miejscu, gdzie tradycyjne obrazy maszyn wirtualnych działają na jednym systemie operacyjnym, obrazy Dockera współdziałają w ramach jednego jądra systemu, przy czym każdy kontener dysponuje swoim indywidualnym systemem plików, co nadaje mu autonomii. Typowo, podstawowy kontener Docker'a charakteryzuje się minimalizmem i lekkością.

Kontenery Dockera, pozostając izolowanymi zarówno względem podstawowego systemu operacyjnego, jak i między sobą, zapewniają znacznie krótszy czas rozruchu aplikacji niż w przypadku tradycyjnej wirtualizacji. Jest to możliwe dzięki ograniczonemu obciążeniu generowanemu przez kontenery.

### 3. Instalacja

#### 3.1 Linux - Ubuntu

Link do oficjalnej instrukcji instalacji : 
https://docs.docker.com/engine/install/ubuntu/

1. Podłączenie repozytorium Dockerowego pod apt.
   
```sh
sudo apt-get update
sudo apt-get install ca-certificates curl gnupg
sudo install -m 0755 -d /etc/apt/keyrings
curl -fsSL https://download.docker.com/linux/ubuntu/gpg | sudo gpg --dearmor -o /etc/apt/keyrings/docker.gpg
sudo chmod a+r /etc/apt/keyrings/docker.gpg

echo \
  "deb [arch="$(dpkg --print-architecture)" signed-by=/etc/apt/keyrings/docker.gpg] https://download.docker.com/linux/ubuntu \
  "$(. /etc/os-release && echo "$VERSION_CODENAME")" stable" | \
  sudo tee /etc/apt/sources.list.d/docker.list > /dev/null
sudo apt-get update
```

2. Instalacja Dockera

```sh
sudo apt-get install docker-ce docker-ce-cli containerd.io docker-buildx-plugin docker-compose-plugin
```

3. Kroki do wykonania po instalacji aby móc konfortowo korzystać z Dockera bez potrzeby używania za każdym razem sudo.

```sh
sudo groupadd docker
sudo usermod -aG docker $USER
```

4. Mamy także możliwość ustawienia Dockera tak aby uruchamiał się zawsze automatycznie po starcie systemu.

```sh
sudo systemctl enable docker.service
sudo systemctl enable containerd.service
```
#### 3.2 Windows

Instrukcja instalacji : 
https://docs.docker.com/desktop/install/windows-install/

Dodatkowo warto mieć na uwadzę w przypadku Windowsa fakt, że Docker jak i wiele innych narzędzi był tworzony z myślą o Linuxie. W związku tym działanie Dockera na Windowsie pozostawia sporo do życzenia, jest powolny i podatny na dziwne błędy. 

Warto w przypadku osób uparcie korzystających z Windowsa zainteresować się projektem od Microsoftu zwanym WSL2. 

https://learn.microsoft.com/en-us/windows/wsl/install

https://learn.microsoft.com/en-us/windows/wsl/install-manual

Jeżeli Państwo zdecydują się na korzystanie z WSL to kolejność instalacji jest najstępująca : 

1. Instalacja WSL koniecznie w wersji 2 !!
2. Instalacja Dockera z zaznaczoną opcją korzystania z wspomnianego wyżej WSL2.

### 4. Podstawowe komendy Dockera

#### 4.1 `docker run`

Polecenie `docker run` służy do uruchomienia kontenera na podstawie konkretnego obrazu. Pozwala także na przekazanie dodatkowych opcji, takich jak mapowanie portów, montowanie woluminów czy definiowanie zmiennych środowiskowych.

```sh
docker run [opcje] obraz [komenda] [argumenty]
```

Przykład:

```sh
docker run -d -p 80:80 nginx
```

#### 4.2 `docker build`

`docker build` tworzy obraz Dockera na podstawie instrukcji zawartych w Dockerfile.

```sh
docker build [opcje] ścieżka
```

Przykład:

```sh
docker build -t moja-aplikacja:latest .
```

#### 4.3 `docker ps`

Polecenie `docker ps` wyświetla informacje o działających kontenerach. Z opcją `-a` pokaże wszystkie kontenery (również te zatrzymane).

```sh
docker ps [-a]
```

#### 4.4 `docker pull`/`docker push`

`docker pull` pobiera obraz z rejestru obrazów.

```sh
docker pull [opcje] nazwa_obrazu[:tag]
```

`docker push` wysyła obraz do rejestru obrazów.

```sh
docker push [opcje] nazwa_obrazu[:tag]
```

Przykład:

```sh
docker pull nginx:latest
docker push moje-repo/moja-aplikacja:latest
```

#### 4.5 `docker commit`

`docker commit` tworzy nowy obraz z istniejącego kontenera, umożliwiając zachowanie aktualnego stanu kontenera jako nowego obrazu.

```sh
docker commit [opcje] kontener nazwa_obrazu[:tag]
```

#### 4.6 `docker exec`

`docker exec` umożliwia wykonywanie poleceń wewnątrz działającego kontenera.

```sh
docker exec [opcje] kontener polecenie [argumenty]
```

#### Inne przydatne polecenia

- `docker stop [kontener]`: zatrzymuje kontener.
- `docker rm [kontener]`: usuwa zatrzymany kontener.
- `docker rmi [obraz]`: usuwa obraz.
- `docker logs [kontener]`: wyświetla logi kontenera.
- `docker network create [nazwa_sieci]`: tworzy nową sieć dla kontenerów.
- `docker volume create [nazwa_woluminu]`: tworzy nowy wolumin.

Pamiętaj, że większość poleceń Docker można używać z różnymi opcjami. Aby uzyskać więcej informacji na temat każdego polecenia i dostępnych dla niego opcji, możesz korzystać z pomocy wbudowanej w CLI Dockera za pomocą `docker [polecenie] --help`.

### 5. Praca z Dockerfile

#### 5.1 Składnia i Struktura Dockerfile

**Dockerfile** to tekstowy plik konfiguracyjny, który zawiera instrukcje budowy obrazu Docker. Składa się z różnych instrukcji, które są wykonane sekwencyjnie.

- `FROM`: Ustanawia obraz bazowy.
- `COPY`: Kopiuje pliki z hosta do obrazu.
- `ADD`: Kopiuje pliki i rozpakowuje archiwa do obrazu.
- `RUN`: Wykonuje polecenie w kontekście budowy obrazu.
- `CMD`: Ustawia domyślne polecenie dla kontenera.
- `ENTRYPOINT`: Ustawia polecenie, które będzie wykonywane w kontenerze.
- `ENV`: Ustawia zmienne środowiskowe.
- `WORKDIR`: Ustawia bieżący katalog roboczy.
- `EXPOSE`: Informuje, że kontener nasłuchuje na określonym porcie.

Przykładowy Dockerfile:

```dockerfile
FROM ubuntu:20.04 
COPY . /app 
WORKDIR /app 
RUN apt-get update && apt-get install -y python3 
CMD ["python3", "app.py"]
```

#### 5.2 Tworzenie Prostego Obrazu Dockera

Biorąc pod uwagę powyższy Dockerfile:

1. Zapisz go w katalogu projektu.
2. Przejdź do katalogu projektu i zbuduj obraz używając:

```sh
docker build -t moja-aplikacja:latest .
```

3. Uruchom kontener na podstawie obrazu:

```sh
docker run -d --name moj_kontener moja-aplikacja:latest
```

#### 5.3 Multistage Builds

**Multistage builds** to technika, która pozwala na użycie wielu etapów budowy obrazu, zmniejszając finalny rozmiar obrazu poprzez eliminowanie niepotrzebnych składników.

Przykład:

```dockerfile
# Etap budowy 
FROM golang:1.17 AS builder 
WORKDIR /src 
COPY . . 
RUN go build -o myapp  

# Etap produkcyjny 
FROM alpine:latest 
COPY --from=builder /src/myapp /myapp 
CMD ["/myapp"]
```

W powyższym Dockerfile budujemy aplikację Go w pierwszym etapie, a następnie kopiujemy tylko skompilowany plik binarny do lekkiego obrazu alpine w drugim etapie.

#### 5.4 Best Practices

- **Używaj konkretnych tagów obrazów bazowych**: Zamiast używania `latest`, używaj konkretnych tagów, aby zapewnić powtarzalność budowy.
- **Minimalizuj liczbę warstw**: Optymalizuj Dockerfile, aby minimalizować liczbę warstw i wielkość obrazu.
- **Unikaj instalacji niepotrzebnych pakietów**: Tylko to, co jest absolutnie konieczne, powinno być obecne w obrazie.
- **Korzystaj z .dockerignore**: Wykorzystaj plik `.dockerignore`, aby wykluczyć niepotrzebne pliki z kontekstu budowy.
- **Optymalizuj kolejność instrukcji**: Umieszczaj często zmieniające się fragmenty kodu jak najniżej w Dockerfile, aby maksymalnie wykorzystać mechanizm cachowania warstw.
- **Używaj wieloetapowych budow**: Aby zminimalizować rozmiar obrazu, używaj techniki multistage builds.

Warto zwrócić uwagę, że to tylko kilka wskazówek. Best practices mogą różnić się w zależności od konkretnego przypadku użycia i technologii, z których korzystasz.

### 6. Docker Compose

#### 6.1 Definicja i Zastosowanie

**Docker Compose** to narzędzie do definiowania i uruchamiania wielokontenerowych aplikacji Docker. Umożliwia użytkownikowi konfigurowanie aplikacji, usług, woluminów i sieci w jednym pliku (`docker-compose.yml`), a następnie uruchomienie wszystkiego jednym poleceniem (`docker-compose up`).

Zastosowania obejmują:

- Lokalny rozwój wielokontenerowych aplikacji.
- Automatyczne testy aplikacji.
- Orkiestrację wielokontenerowej aplikacji na środowisku produkcyjnym.

#### 6.2 Struktura Pliku `docker-compose.yml`

Podstawowe elementy pliku `docker-compose.yml` to:

- `version`: Wersja składni pliku Docker Compose.
- `services`: Definicje kontenerów, które mają być uruchomione.
- `networks`: Definicje sieci, które mają być utworzone.
- `volumes`: Definicje woluminów, które mają być utworzone.

Przykładowa struktura:

```yaml
version: '3'

services:
  web:
    image: nginx:latest
    ports:
      - "8080:80"
  db:
    image: postgres:latest
    environment:
      POSTGRES_DB: exampledb
      POSTGRES_USER: user
      POSTGRES_PASSWORD: password

networks:
  example-network:
    driver: bridge

volumes:
  example-volume:

```

#### 6.3 Przykłady Użycia i Orkiestracji Wielokontenerowych Aplikacji

Przykład: Aplikacja webowa i baza danych

Przyjmując, że mamy aplikację webową zbudowaną na nginx (serwer web) i korzystamy z PostgreSQL jako bazy danych, możemy zdefiniować następujący `docker-compose.yml`:

```yaml
version: '3'

services:
  web:
    image: nginx:latest
    ports:
      - "8080:80"
  db:
    image: postgres:latest
    environment:
      POSTGRES_DB: exampledb
      POSTGRES_USER: user
      POSTGRES_PASSWORD: password
```

- `web`: usługa, która uruchamia serwer nginx i mapuje port 8080 hosta na port 80 kontenera.
- `db`: usługa, która uruchamia PostgreSQL i ustawia zmienne środowiskowe, aby zainicjować bazę danych przy pierwszym uruchomieniu.

Aby uruchomić powyższą konfigurację, należy użyć polecenia:

```sh
docker-compose up
```

To uruchomi obie usługi w tle. Jeśli chcemy je uruchomić w trybie pierwszego planu (czyli w konsoli, w której uruchomiliśmy polecenie), możemy użyć:

```sh
docker-compose up -d
```

Inne przydatne polecenia `docker-compose` to między innymi:

- `docker-compose down`: Zatrzymuje i usuwa kontenery, sieci i woluminy.
- `docker-compose ps`: Wyświetla listę uruchomionych kontenerów oraz ich status.
- `docker-compose logs`: Wyświetla logi kontenerów.

Oczywiście, Docker Compose jest narzędziem o wiele bardziej rozbudowanym i posiada wiele innych opcji, które mogą być przydatne w różnych scenariuszach użycia.




