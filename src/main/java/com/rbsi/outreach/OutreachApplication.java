package com.rbsi.outreach;

import com.rbsi.outreach.service.GenerateLetterService;
import com.rbsi.outreach.service.SendEmailService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Bean;

@SpringBootApplication
public class OutreachApplication {

	public static void main(String[] args) {
		SpringApplication.run(OutreachApplication.class, args);
	}

	@Autowired
	private GenerateLetterService generateLetterService;
//	@Bean
//	public CommandLineRunner run(GenerateLetterService generateLetterService, SendEmailService sendEmailService) {
//		return args -> {
//			if (args.length > 0) {
//				String serviceToRun = args[0];
//
//				switch (serviceToRun.toLowerCase()) {
//					case "generate":
//						System.out.println("Running GenerateLetterService...");
//						generateLetterService.generateLetters();
//						break;
//
//					case "send":
//						System.out.println("Running SendEmailService...");
//						sendEmailService.sendEmails();
//						break;
//
//					default:
//						System.out.println("Invalid argument. Use 'generate' or 'send'.");
//				}
//			} else {
//				System.out.println("No argument provided. Use 'generate' or 'send' to run a specific service.");
//			}
//		};
//	}

	@Bean
	public CommandLineRunner run() {
		return args -> {
			System.out.println("Running Generate Letter Flow...");
			generateLetterService.generateLetters();
			System.out.println("Generate Letter Flow Completed");
		};
	}
}