plugins {
    `java`
    `application`
    `maven-publish`
}

configure<JavaApplication> {
    mainClass.set("com.github.emotionbug.mmdtopptx.MermaidSvg2Pptx")
}

group = "com.github.emotionbug"
version = "1.0.0"
description = "mermaid-to-pptx"

repositories {
    mavenCentral()
}

dependencies {
    implementation("org.apache.logging.log4j:log4j-api:2.25.3")
    implementation("org.apache.logging.log4j:log4j-core:2.25.3")
    implementation("org.apache.logging.log4j:log4j-slf4j2-impl:2.25.3")

    implementation("org.apache.poi:poi-ooxml:5.5.1")

    implementation("org.seleniumhq.selenium:selenium-java:4.39.0")
    implementation("io.github.bonigarcia:webdrivermanager:6.3.3")
}

java {
    toolchain {
        languageVersion.set(JavaLanguageVersion.of(21))
    }
}

publishing {
    publications.create<MavenPublication>("maven") {
        from(components["java"])
    }
}

tasks.withType<JavaCompile>().configureEach {
    options.encoding = "UTF-8"
}

tasks.withType<Javadoc>().configureEach {
    options.encoding = "UTF-8"
}
