plugins {
    id("groovy")
}

group = "ru.kazantsev"
version = "1.0-SNAPSHOT"

repositories {
    mavenCentral()
}

dependencies {
    implementation("org.codehaus.groovy:groovy-all:3.0.19")
    testImplementation(platform("org.junit:junit-bom:5.9.1"))
    testImplementation("org.junit.jupiter:junit-jupiter")
    implementation("org.apache.poi:poi-ooxml:5.2.4")
}

tasks.test {
    useJUnitPlatform()
}
