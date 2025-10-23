ifdef CI
	DOCKER_PHP_EXEC :=
	DOCKER_BUILD :=
else
	DOCKER_PHP_EXEC := docker compose run --rm php
	DOCKER_BUILD := docker compose build --pull
endif
PHP_BIN=php -d zend.assertions=1

SRCS := $(shell find ./lib ./tests -type f -not -path "*/tmp/*")

LOCAL_BASE_BRANCH ?= $(shell git show-branch | sed "s/].*//" | grep "\*" | grep -v "$$(git rev-parse --abbrev-ref HEAD)" | head -n1 | sed "s/^.*\[//")
ifeq ($(strip $(LOCAL_BASE_BRANCH)),)
	LOCAL_BASE_BRANCH := HEAD^
endif
BASE_BRANCH ?= $(LOCAL_BASE_BRANCH)

all: csfix static-analysis code-coverage
	@echo "Done."

.env: /etc/passwd /etc/group Makefile
	printf "USER_ID=%s\nGROUP_ID=%s\n" `id --user "${USER}"` `id --group "${USER}"` > .env

vendor: .env docker-compose.yml Dockerfile composer.json
	$(DOCKER_BUILD)
	$(DOCKER_PHP_EXEC) composer update
	$(DOCKER_PHP_EXEC) composer bump
	touch --no-create $@

.PHONY: csfix
csfix: vendor
	$(DOCKER_PHP_EXEC) vendor/bin/php-cs-fixer fix -v

.PHONY: static-analysis
static-analysis: vendor
	$(DOCKER_PHP_EXEC) $(PHP_BIN) vendor/bin/phpstan analyse --memory-limit=512M $(PHPSTAN_ARGS)

coverage/junit.xml: vendor $(SRCS) Makefile
	$(DOCKER_PHP_EXEC) $(PHP_BIN) vendor/bin/phpunit $(PHPUNIT_ARGS)

.PHONY: test
test: coverage/junit.xml

.PHONY: code-coverage
code-coverage: coverage/junit.xml
	echo "Base branch: $(BASE_BRANCH)"
	$(DOCKER_PHP_EXEC) $(PHP_BIN) \
		vendor/bin/infection \
		--threads=$(shell nproc) \
		--git-diff-lines \
		--git-diff-base=$(BASE_BRANCH) \
		--skip-initial-tests \
		--coverage=coverage \
		--show-mutations \
		--verbose \
		--min-msi=100 \
		$(INFECTION_ARGS)
