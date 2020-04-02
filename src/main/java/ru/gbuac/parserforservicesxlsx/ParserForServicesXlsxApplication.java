package ru.gbuac.parserforservicesxlsx;

import lombok.extern.slf4j.Slf4j;
import org.springframework.boot.ApplicationArguments;
import org.springframework.boot.ApplicationRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import ru.gbuac.parserforservicesxlsx.service.ParserService;

@Slf4j
@SpringBootApplication
public class ParserForServicesXlsxApplication implements ApplicationRunner {

    private final ParserService parserService;

    public ParserForServicesXlsxApplication(ParserService parserService) {
        this.parserService = parserService;
    }

    public static void main(String[] args) {
        SpringApplication.run(ParserForServicesXlsxApplication.class, args);
    }

    @Override
    public void run(ApplicationArguments args) {
        try {
            parserService.startProcess();
        } catch (Exception e) {
            log.error(e.getMessage(), e);
        }
    }

}